import os
import io
import json
import time
import uuid
import tempfile
from pathlib import Path
from typing import Dict, List, Tuple, Optional

import streamlit as st
import fitz  # PyMuPDF
import google.generativeai as genai
from langchain.text_splitter import RecursiveCharacterTextSplitter
from langchain_community.document_loaders import PyMuPDFLoader
from pptx import Presentation
from pptx.util import Inches
from dotenv import load_dotenv
import requests
from PIL import Image
import re

# Load environment variables
load_dotenv()

# Set page configuration
st.set_page_config(
    page_title="PDF to PowerPoint Generator",
    page_icon="📊",
    layout="wide"
)

# Constants
TEMP_DIR = Path("temp_files")
IMAGE_DIR = TEMP_DIR / "images"
OUTPUT_DIR = Path("output")

# Create necessary directories
TEMP_DIR.mkdir(exist_ok=True)
IMAGE_DIR.mkdir(exist_ok=True)
OUTPUT_DIR.mkdir(exist_ok=True)

# Function to setup API key
def setup_gemini_api():
    api_key = st.session_state.get("api_key", os.getenv("GEMINI_API_KEY", ""))
    
    if not api_key:
        st.warning("Gemini API key is not set. Please enter it below.")
        return False
    
    try:
        genai.configure(api_key=api_key)
        # Test connection with updated model name
        model = genai.GenerativeModel('gemini-1.5-flash')
        return True
    except Exception as e:
        st.error(f"Error configuring Gemini API: {str(e)}")
        return False

# Function to extract images from PDF
def extract_images_from_pdf(pdf_path: str) -> Dict[int, List[str]]:
    """Extract images from PDF and save them to disk"""
    image_paths_by_page = {}
    
    with st.spinner("Extracting images from PDF..."):
        try:
            pdf_document = fitz.open(pdf_path)
            
            for page_num, page in enumerate(pdf_document):
                image_list = page.get_images(full=True)
                
                if not image_list:
                    continue
                
                image_paths_by_page[page_num] = []
                
                for img_index, img_info in enumerate(image_list):
                    xref = img_info[0]
                    base_image = pdf_document.extract_image(xref)
                    image_bytes = base_image["image"]
                    
                    # Save image
                    image_filename = f"page_{page_num+1}_img_{img_index+1}.png"
                    image_path = str(IMAGE_DIR / image_filename)
                    
                    with open(image_path, "wb") as img_file:
                        img_file.write(image_bytes)
                    
                    image_paths_by_page[page_num].append(image_path)
                
                # Update progress
                progress_bar.progress((page_num + 1) / len(pdf_document))
            
            return image_paths_by_page
        
        except Exception as e:
            st.error(f"Error extracting images: {str(e)}")
            return {}

# Function to get text content from PDF pages
def extract_text_from_pdf(pdf_path: str) -> Dict[int, str]:
    """Extract text content from each page in the PDF"""
    text_by_page = {}
    
    with st.spinner("Extracting text from PDF..."):
        try:
            loader = PyMuPDFLoader(pdf_path)
            pages = loader.load_and_split()
            
            for page in pages:
                page_num = page.metadata.get("page", 0)
                text_by_page[page_num] = page.page_content
            
            return text_by_page
        
        except Exception as e:
            st.error(f"Error extracting text: {str(e)}")
            return {}

# Function to generate image captions using Gemini
def generate_image_captions(image_paths_by_page: Dict[int, List[str]], 
                           text_by_page: Dict[int, str]) -> Dict[str, Dict]:
    """Generate captions for images using Gemini API"""
    captions_data = {}
    
    with st.spinner("Generating image captions with Gemini..."):
        # Update model name here
        model = genai.GenerativeModel('gemini-1.5-flash')
        
        total_pages = len(image_paths_by_page)
        current_page = 0
        
        for page_num, image_paths in image_paths_by_page.items():
            page_text = text_by_page.get(page_num, "")
            
            for img_path in image_paths:
                # Rate limiting
                time.sleep(1)  # Simple rate limiting
                
                try:
                    img = Image.open(img_path)
                    
                    prompt = f"""
                    I need a detailed caption for this image.
                    The image appears on page {page_num+1} of a document.
                    Here's the text context from this page:
                    {page_text[:500]}...
                    
                    Provide a JSON response with:
                    1. "caption": A descriptive caption
                    2. "alt_text": Short alternative text
                    3. "relevance": Brief note on relevance to the document
                    """
                    
                    response = model.generate_content([prompt, img])
                    response_text = response.text
                    
                    # Extract JSON from response
                    json_match = re.search(r'```json\n(.*?)\n```', response_text, re.DOTALL)
                    if json_match:
                        caption_json = json.loads(json_match.group(1))
                    else:
                        # Try to extract JSON without code block markers
                        json_match = re.search(r'(\{.*\})', response_text, re.DOTALL)
                        if json_match:
                            caption_json = json.loads(json_match.group(1))
                        else:
                            caption_json = {
                                "caption": "Could not generate caption",
                                "alt_text": "Image",
                                "relevance": "Unknown"
                            }
                    
                    img_name = os.path.basename(img_path)
                    captions_data[img_path] = caption_json
                    
                except Exception as e:
                    st.warning(f"Error processing image {img_path}: {str(e)}")
                    captions_data[img_path] = {
                        "caption": "Error generating caption",
                        "alt_text": "Image",
                        "relevance": "Error in processing"
                    }
            
            current_page += 1
            progress_bar.progress(current_page / total_pages)
    
    # Save captions to JSON file
    with open(TEMP_DIR / "image_captions.json", "w") as f:
        json.dump(captions_data, f, indent=2)
    
    return captions_data


# Function to validate uploaded PowerPoint template
def validate_pptx_template(uploaded_file):
    """Validate that the uploaded file is a valid PowerPoint template"""
    try:
        # Save the uploaded file temporarily
        template_path = TEMP_DIR / "uploaded_template.pptx"
        with open(template_path, "wb") as f:
            f.write(uploaded_file.read())
        
        # Try to open it with python-pptx to validate
        prs = Presentation(template_path)
        
        # Check if it has at least some slide layouts
        if len(prs.slide_layouts) < 1:
            return False, "Template has no slide layouts"
        
        # Instead of requiring title in first layout, check if ANY layout has a title placeholder
        has_title_layout = False
        for layout in prs.slide_layouts:
            if any(shape.placeholder_format.type == 1 for shape in layout.placeholders):
                has_title_layout = True
                break
        
        if not has_title_layout:
            return False, "Template doesn't have any layouts with title placeholders"
        
        # Check if template has at least one content placeholder in any layout
        has_content_layout = False
        for layout in prs.slide_layouts:
            for shape in layout.placeholders:
                if shape.placeholder_format.type in [2, 5, 7, 17]:  # Common content placeholder types
                    has_content_layout = True
                    break
            if has_content_layout:
                break
        
        if not has_content_layout:
            return False, "Template doesn't have any layouts with content placeholders"
        
        # Return success with the template path
        return True, str(template_path)
    
    except Exception as e:
        st.warning(f"Template validation warning: {str(e)}")
        # Still return the template even if validation has issues - more permissive approach
        return True, str(template_path)

# Function to optimize prompt based on presentation time
def optimize_prompt_for_presentation_time(basic_prompt, pdf_summary, presentation_time):
    """Use multi-step prompt optimization based on presentation time"""
    with st.spinner("Optimizing presentation structure..."):
        try:
            model = genai.GenerativeModel('gemini-1.5-flash')
            
            # First step: Analyze content and time requirements
            system_prompt = """
            You are an expert presentation coach specializing in time management for presentations.
            Based on the presentation topic and available time, determine:
            1. Optimal number of slides
            2. Content density per slide
            3. Time allocation strategy
            
            Provide specific, actionable guidance on how to structure this presentation.
            """
            
            time_analysis_prompt = f"""
            Basic presentation request: {basic_prompt}
            
            Document summary: {pdf_summary[:500]}...
            
            Available presentation time: {presentation_time} minutes
            
            Please analyze and provide structured guidance for this presentation.
            """
            
            # Rate limiting
            time.sleep(1)
            
            response = model.generate_content(
                [system_prompt, time_analysis_prompt],
                generation_config={"temperature": 0.2}
            )
            
            time_analysis = response.text
            
            # Second step: Create optimized prompt with presentation structure
            system_prompt_2 = """
            You are an expert presentation designer who optimizes content for specific time constraints.
            Convert the provided analysis into a detailed, structured prompt that will guide the creation 
            of a well-paced presentation matching the time requirements.
            
            Your output should be a comprehensive prompt that specifies:
            1. Exact number of slides to create
            2. Specific slide types and their order
            3. Content density guidelines for each section
            4. Transitions and timing suggestions
            """
            
            optimization_prompt = f"""
            Original request: {basic_prompt}
            
            Time analysis: {time_analysis}
            
            Available presentation time: {presentation_time} minutes
            
            Document summary: {pdf_summary[:500]}...
            
            Create an optimized, structured prompt for generating this presentation.
            """
            
            # Rate limiting
            time.sleep(1)
            
            response = model.generate_content(
                [system_prompt_2, optimization_prompt],
                generation_config={"temperature": 0.2}
            )
            
            optimized_prompt = response.text
            
            # Combine with original prompt for final result
            final_prompt = f"""
            {basic_prompt}
            
            PRESENTATION TIME: {presentation_time} minutes
            
            STRUCTURAL GUIDANCE:
            {optimized_prompt}
            """
            
            return final_prompt
            
        except Exception as e:
            st.warning(f"Error optimizing prompt: {str(e)}")
            return basic_prompt  # Fallback to original prompt if optimization fails

# Update the generate_presentation_content function to better adapt to templates
def generate_presentation_content(pdf_path: str, captions_data: Dict, user_prompt: str, template_path: Optional[str] = None) -> Dict:
    """Generate presentation content using Gemini with template adaptation"""
    with st.spinner("Generating presentation content..."):
        try:
            # Extract sample text from PDF for context
            loader = PyMuPDFLoader(pdf_path)
            pages = loader.load_and_split()
            
            all_text = "\n\n".join([page.page_content for page in pages])
         
            # Create a comprehensive summary of the full document
            full_text_summary = all_text[:100000] if len(all_text) > 10000 else all_text
            
            # Prepare image references for the model
            image_references = []
            for img_path, caption_info in captions_data.items():
                image_references.append({
                    "path": img_path,
                    "caption": caption_info["caption"],
                    "alt_text": caption_info["alt_text"]
                })
            
            # Analyze template if provided to extract available layouts
            template_guidance = ""
            if template_path and os.path.exists(template_path):
                try:
                    prs = Presentation(template_path)
                    layouts_info = []
                    
                    for i, layout in enumerate(prs.slide_layouts):
                        placeholder_types = []
                        for shape in layout.placeholders:
                            placeholder_types.append(f"Type {shape.placeholder_format.type}")
                        
                        layouts_info.append(f"Layout {i}: {layout.name} - Placeholders: {', '.join(placeholder_types)}")
                    
                    template_guidance = f"""
                    TEMPLATE INFORMATION:
                    This presentation should be adapted to the following template layouts:
                    {chr(10).join(layouts_info)}
                    
                    Ensure content is structured to match these layouts.
                    """
                except Exception as e:
                    template_guidance = "Note: Template analysis failed. Create standard slide formats."
            
            # Prepare prompt for Gemini with template guidance
            model = genai.GenerativeModel('gemini-1.5-pro')  # Use higher quality model
            
            system_prompt = """
            You are an expert presentation designer specialized in creating highly structured, 
            professional PowerPoint presentations that fit precisely into templates.
            
            Create a presentation with the following slide structure:
            1. Title Slide: Title, subtitle
            2. Agenda/Overview: List of key topics
            3. Content Slides: Each with clear title, 3-5 bullet points or short paragraphs, and relevant images
            4. Summary/Conclusion: Recap of key points
            
            For each slide, ensure:
            - Clear hierarchical structure (main title, subtitle if needed, content)
            - Logical flow from previous slide
            - Concise, scannable content (not walls of text)
            - Professional tone and language
            - Appropriate image placement
            
            IMPORTANT: Your response must be valid JSON with no syntax errors. Double-check all commas, 
            brackets, and ensure all strings are properly quoted. Do not include any explanatory text 
            outside the JSON structure.
            
            Generate a JSON structure with:
            {
                "title": "Main Presentation Title",
                "subtitle": "Optional Subtitle",
                "slides": [
                    {
                        "slide_type": "title_slide|section_slide|content_slide|image_slide|conclusion_slide",
                        "title": "Slide Title",
                        "subtitle": "Optional Slide Subtitle",
                        "content": ["Point 1", "Point 2", "Point 3"],
                        "notes": "Speaker notes for this slide",
                        "images": ["image_path_1", "image_path_2"]
                    }
                ]
            }
            """
            
            user_message = f"""
            {user_prompt}
            
            PDF CONTENT (comprehensive extract from the document):
            {full_text_summary}
            
            NOTE: The above represents a substantial portion of the full document content. 
            Your task is to create a presentation based on this content, capturing the key 
            themes, findings, and important points from the entire document.
            
            Available images (reference these by path when creating slides):
            {json.dumps(image_references, indent=2)}
            
            {template_guidance}
            
            Create a cohesive presentation with logical flow, ensuring:
            1. Clear transitions between topics
            2. Consistent structure on each slide (title + content)
            3. Proper distribution of content (not too dense, not too sparse)
            4. Strategic use of available images where relevant
            5. Professional, engaging language
            
            Include a title slide, agenda slide, and conclusion slide in your structure.
            
            IMPORTANT: Your response must be valid, properly formatted JSON with no syntax errors.
            Ensure all arrays and objects have proper delimiters (commas between items).
            """
            
            # Rate limiting
            time.sleep(2)
            
            response = model.generate_content(
                [system_prompt, user_message],
                generation_config={"temperature": 0.2, "max_output_tokens": 8192}
            )
            
            # Extract JSON from response
            response_text = response.text
            
            # First try to extract JSON with code block markers
            json_match = re.search(r'```json\n(.*?)\n```', response_text, re.DOTALL)
            
            if json_match:
                json_str = json_match.group(1)
            else:
                # Try to extract JSON without code block markers
                json_match = re.search(r'(\{.*\})', response_text, re.DOTALL)
                if json_match:
                    json_str = json_match.group(1)
                else:
                    raise ValueError("Could not extract JSON from response")
            
            # Try to parse the JSON with additional error handling
            try:
                presentation_content = json.loads(json_str)
            except json.JSONDecodeError as e:
                st.warning(f"Initial JSON parsing failed: {str(e)}")
                
                # Log the problematic JSON for debugging
                with open(TEMP_DIR / "json_error.txt", "w") as f:
                    f.write(f"Error: {str(e)}\n\n")
                    f.write(json_str)
                
                # Try to fix common JSON syntax errors and retry
                fixed_json = fix_json_syntax(json_str)
                try:
                    presentation_content = json.loads(fixed_json)
                    st.info("Fixed JSON syntax errors automatically")
                except json.JSONDecodeError:
                    # If still failing, request a cleaner JSON response
                    st.warning("Retrying with explicit request for clean JSON...")
                    
                    repair_prompt = """
                    The JSON you provided has syntax errors. Please generate a corrected version
                    with proper JSON syntax. Common issues include:
                    1. Missing commas between array or object items
                    2. Trailing commas at the end of arrays or objects
                    3. Unquoted property names
                    4. Mismatched quotes or brackets
                    
                    Respond ONLY with the corrected JSON structure, nothing else.
                    """
                    
                    # Rate limiting
                    time.sleep(2)
                    
                    repair_response = model.generate_content([repair_prompt, json_str])
                    repair_text = repair_response.text
                    
                    # Try to extract and parse the repaired JSON
                    json_match = re.search(r'```json\n(.*?)\n```', repair_text, re.DOTALL)
                    if json_match:
                        repaired_json = json_match.group(1)
                    else:
                        repaired_json = re.search(r'(\{.*\})', repair_text, re.DOTALL).group(1)
                    
                    presentation_content = json.loads(repaired_json)
            
            # Save presentation content
            with open(TEMP_DIR / "presentation_content.json", "w") as f:
                json.dump(presentation_content, f, indent=2)
            
            return presentation_content
            
        except Exception as e:
            st.error(f"Error generating presentation content: {str(e)}")
            return {}

def fix_json_syntax(json_str):
    """Attempt to fix common JSON syntax errors"""
    # Fix missing commas between objects in arrays
    fixed = re.sub(r'}\s*{', '},{', json_str)
    
    # Fix missing commas between array elements
    fixed = re.sub(r'"\s*"', '","', fixed)
    
    # Fix trailing commas in arrays and objects
    fixed = re.sub(r',\s*}', '}', fixed)
    fixed = re.sub(r',\s*\]', ']', fixed)
    
    # Fix unquoted property names
    fixed = re.sub(r'(\w+):', r'"\1":', fixed)
    
    # Remove any extra text before/after the JSON object
    json_obj_match = re.search(r'({.*})', fixed, re.DOTALL)
    if json_obj_match:
        fixed = json_obj_match.group(1)
    
    return fixed

# Update the create_powerpoint function for better template compatibility
def create_powerpoint(presentation_content: Dict, template_path: Optional[str] = None) -> str:
    """Create PowerPoint file from content with improved template handling"""
    with st.spinner("Creating PowerPoint presentation..."):
        try:
            # Create presentation with template if provided
            if template_path and os.path.exists(template_path):
                prs = Presentation(template_path)
            else:
                prs = Presentation()
            
            # Create a comprehensive map of available layouts and their capabilities
            layout_map = analyze_template_layouts(prs)
            
            # Add title slide
            add_title_slide(prs, layout_map, presentation_content)
            
            # Add content slides with appropriate layout selection
            for slide_content in presentation_content.get("slides", []):
                add_content_slide(prs, layout_map, slide_content)
            
            # Save the presentation with timestamp
            timestamp = time.strftime("%Y%m%d_%H%M%S")
            output_filename = f"presentation_{timestamp}.pptx"
            output_path = str(OUTPUT_DIR / output_filename)
            prs.save(output_path)
            
            return output_path
            
        except Exception as e:
            st.error(f"Error creating PowerPoint: {str(e)}")
            return ""

def analyze_template_layouts(prs: Presentation) -> Dict:
    """Analyze template layouts and create a detailed map of available layouts and their capabilities"""
    layout_map = {
        "title_layouts": [],
        "content_layouts": [],
        "section_layouts": [],
        "image_layouts": [],
        "multi_content_layouts": [],
        "list_layouts": [],
        "blank_layouts": [],
        "all_layouts": []
    }
    
    for idx, layout in enumerate(prs.slide_layouts):
        layout_info = {
            "index": idx,
            "name": layout.name.lower() if hasattr(layout, 'name') else f"Layout {idx}",
            "placeholder_types": {},
            "has_title": False,
            "has_content": False,
            "has_image": False,
            "has_list": False,
            "has_chart": False,
            "has_table": False,
            "has_subtitle": False
        }
        
        # Analyze placeholders
        for shape in layout.placeholders:
            ph_type = shape.placeholder_format.type
            ph_name = shape.name.lower() if hasattr(shape, 'name') else f"Type {ph_type}"
            
            layout_info["placeholder_types"][ph_type] = ph_name
            
            # Detect capabilities based on placeholder types
            if ph_type == 1:  # Title
                layout_info["has_title"] = True
            elif ph_type == 2:  # Subtitle/body
                layout_info["has_subtitle"] = True
            elif ph_type in [7, 17]:  # Content, body text
                layout_info["has_content"] = True
                # Check if it likely supports bullets/lists
                if hasattr(shape, 'text_frame') and hasattr(shape.text_frame, 'paragraphs'):
                    layout_info["has_list"] = True
            elif ph_type in [18, 19]:  # Picture, media
                layout_info["has_image"] = True
            elif ph_type == 3:  # Chart
                layout_info["has_chart"] = True
            elif ph_type == 4:  # Table
                layout_info["has_table"] = True
        
        # Categorize layout based on capabilities
        layout_map["all_layouts"].append(layout_info)
        
        name = layout_info["name"]
        
        # Classify layouts based on name and capabilities
        if "title" in name and not any(x in name for x in ["content", "body"]):
            layout_map["title_layouts"].append(layout_info)
        elif "section" in name or "divider" in name:
            layout_map["section_layouts"].append(layout_info)
        elif layout_info["has_content"] and layout_info["has_image"]:
            layout_map["image_layouts"].append(layout_info)
        elif layout_info["has_content"]:
            if "list" in name or layout_info["has_list"]:
                layout_map["list_layouts"].append(layout_info)
            layout_map["content_layouts"].append(layout_info)
        elif "blank" in name:
            layout_map["blank_layouts"].append(layout_info)
        elif layout_info["has_image"]:
            layout_map["image_layouts"].append(layout_info)
        
        # Also check for multi-content layouts
        if layout_info["has_content"] and len([k for k, v in layout_info["placeholder_types"].items() 
                                              if k in [7, 17]]) > 1:
            layout_map["multi_content_layouts"].append(layout_info)
    
    return layout_map

def add_title_slide(prs: Presentation, layout_map: Dict, presentation_content: Dict):
    """Add the title slide using the most appropriate layout"""
    # Find the best title slide layout
    title_layouts = layout_map["title_layouts"]
    
    if title_layouts:
        # Prefer layouts with both title and subtitle if we have both
        if "subtitle" in presentation_content and presentation_content["subtitle"]:
            suitable_layouts = [l for l in title_layouts if l["has_subtitle"]]
            layout_idx = suitable_layouts[0]["index"] if suitable_layouts else title_layouts[0]["index"]
        else:
            layout_idx = title_layouts[0]["index"]
    else:
        # Fallback to first layout
        layout_idx = 0
    
    slide = prs.slides.add_slide(prs.slide_layouts[layout_idx])
    
    # Add title and subtitle where appropriate
    for shape in slide.placeholders:
        ph_type = shape.placeholder_format.type
        
        if ph_type == 1:  # Title
            shape.text = presentation_content.get("title", "Generated Presentation")
        elif ph_type == 2:  # Subtitle
            shape.text = presentation_content.get("subtitle", "")

def select_layout_for_content(layout_map: Dict, slide_content: Dict) -> int:
    """Select the most appropriate layout based on content characteristics"""
    slide_type = slide_content.get("slide_type", "content_slide")
    has_images = "images" in slide_content and slide_content["images"]
    has_list = "content" in slide_content and isinstance(slide_content["content"], list) and len(slide_content["content"]) > 1
    
    # Map content characteristics to layout types
    if slide_type == "title_slide":
        layouts = layout_map["title_layouts"]
    elif slide_type == "section_slide":
        layouts = layout_map["section_layouts"]
    elif slide_type == "image_slide" or (has_images and not has_list):
        layouts = layout_map["image_layouts"]
    elif slide_type == "content_slide":
        if has_list:
            # Try to find list layouts first
            layouts = layout_map["list_layouts"] if layout_map["list_layouts"] else layout_map["content_layouts"]
        else:
            layouts = layout_map["content_layouts"]
    else:
        layouts = layout_map["content_layouts"]
    
    # If we found suitable layouts, use the first one
    if layouts:
        return layouts[0]["index"]
    
    # If no matching layouts, choose based on content type
    if has_images:
        # Try to find any layout with image support
        for layout_info in layout_map["all_layouts"]:
            if layout_info["has_image"]:
                return layout_info["index"]
    
    # Default to first content layout if available, otherwise first layout
    if layout_map["content_layouts"]:
        return layout_map["content_layouts"][0]["index"]
    return 0

def add_content_slide(prs: Presentation, layout_map: Dict, slide_content: Dict):
    """Add a content slide with intelligent layout selection and content placement"""
    # Select the most appropriate layout for this content
    layout_idx = select_layout_for_content(layout_map, slide_content)
    slide = prs.slides.add_slide(prs.slide_layouts[layout_idx])
    
    # Get layout information for this slide
    layout_info = next((l for l in layout_map["all_layouts"] if l["index"] == layout_idx), None)
    
    # Add slide title if there's a title placeholder
    title_shape = None
    for shape in slide.placeholders:
        if shape.placeholder_format.type == 1:  # Title
            shape.text = slide_content.get("title", "")
            title_shape = shape
            break
    
    # Handle content (bullet points, text, etc.)
    if "content" in slide_content and slide_content["content"]:
        content_added = False
        
        # First try to find dedicated content placeholders
        content_shapes = [s for s in slide.placeholders 
                         if s.placeholder_format.type in [7, 17]]  # Content/body
        
        if content_shapes:
            shape = content_shapes[0]  # Use first content shape
            text_frame = shape.text_frame
            text_frame.clear()
            
            # Add content as paragraphs/bullets
            for i, paragraph_text in enumerate(slide_content["content"]):
                if i == 0:
                    p = text_frame.paragraphs[0]
                else:
                    p = text_frame.add_paragraph()
                p.text = paragraph_text
                p.level = 0  # Top level bullet
            
            content_added = True
        
        # If no dedicated content placeholder, try other text placeholders
        if not content_added:
            # Try subtitle placeholder
            for shape in slide.placeholders:
                if shape.placeholder_format.type == 2 and shape != title_shape:  # Subtitle, not the same as title
                    text_frame = shape.text_frame
                    text_frame.clear()
                    
                    # Join content with line breaks if more than one item
                    if isinstance(slide_content["content"], list):
                        shape.text = "\n".join(slide_content["content"])
                    else:
                        shape.text = slide_content["content"]
                    
                    content_added = True
                    break
        
        # Last resort: add text box if no suitable placeholder found
        if not content_added:
            left = Inches(1)
            top = Inches(2.5)
            width = Inches(8)
            height = Inches(4)
            
            # Position below title if title exists
            if title_shape:
                title_top = title_shape.top
                title_height = title_shape.height
                top = title_top + title_height + Inches(0.2)
            
            txBox = slide.shapes.add_textbox(left, top, width, height)
            tf = txBox.text_frame
            
            # Add content as paragraphs
            for i, paragraph_text in enumerate(slide_content["content"]):
                if i == 0:
                    p = tf.paragraphs[0]
                else:
                    p = tf.add_paragraph()
                p.text = paragraph_text
    
    # Handle images
    if "images" in slide_content and slide_content["images"]:
        # First try to find image placeholders
        image_placeholders = [s for s in slide.placeholders 
                             if s.placeholder_format.type in [18, 19]]  # Picture
        
        images = slide_content["images"]
        
        if image_placeholders and images:
            # Try to match images to placeholders
            for i, img_path in enumerate(images):
                if i < len(image_placeholders) and os.path.exists(img_path):
                    try:
                        image_placeholders[i].insert_picture(img_path)
                    except Exception:
                        # Fallback if insert_picture fails
                        placeholder = image_placeholders[i]
                        left = placeholder.left
                        top = placeholder.top
                        width = placeholder.width
                        height = placeholder.height
                        slide.shapes.add_picture(img_path, left, top, width, height)
        else:
            # No placeholders or too many images - add directly
            content_shapes = [s for s in slide.placeholders 
                              if s.placeholder_format.type in [7, 17]]
            
            # Calculate positions based on layout
            for i, img_path in enumerate(images):
                if os.path.exists(img_path):
                    if len(images) == 1:
                        # For single image, place it centered
                        if layout_info and layout_info["has_content"] and content_shapes:
                            # If content area exists, place below it
                            content_shape = content_shapes[0]
                            left = content_shape.left
                            top = content_shape.top + content_shape.height + Inches(0.5)
                            width = content_shape.width
                        else:
                            # Center in slide
                            left = Inches(2)
                            top = Inches(3)
                            width = Inches(6)
                    else:
                        # For multiple images, arrange in grid
                        row = i // 2
                        col = i % 2
                        left = Inches(1 + col * 4.25)
                        top = Inches(3 + row * 2.5)
                        width = Inches(4)
                    
                    # Add the image
                    slide.shapes.add_picture(img_path, left, top, width=width)
    
    # Add notes if present
    if "notes" in slide_content:
        notes_slide = slide.notes_slide
        text_frame = notes_slide.notes_text_frame
        text_frame.text = slide_content["notes"]

# UI Components
st.title("PDF to PowerPoint Generator")
st.markdown("Upload a PDF and generate a PowerPoint presentation with AI assistance.")

# API Key input
with st.expander("API Key Configuration", expanded=not os.getenv("GEMINI_API_KEY")):
    api_key_input = st.text_input(
        "Gemini API Key",
        value=os.getenv("GEMINI_API_KEY", ""),
        type="password",
        key="api_key"
    )
    
    if st.button("Test API Connection"):
        if setup_gemini_api():
            st.success("API connection successful!")
        else:
            st.error("Failed to connect to API.")

# File upload
uploaded_file = st.file_uploader("Upload PDF", type=["pdf"])

# Presentation configuration
with st.expander("Presentation Settings", expanded=True):
    user_prompt = st.text_area(
        "Describe what you want in the presentation",
        "Create a concise, professional presentation that summarizes the key points in the document."
    )
    
    # Add presentation time input
    presentation_time = st.number_input(
        "Presentation Duration (minutes)",
        min_value=5,
        max_value=120,
        value=15,
        step=5,
        help="Specify how long the presentation will be. This helps optimize content amount and slide count."
    )
    
    # Template upload option
    template_file = st.file_uploader("Upload PowerPoint Template (Optional)", type=["pptx", "potx"])
    template_path = None
    
    if template_file:
        valid_template, template_result = validate_pptx_template(template_file)
        if valid_template:
            template_path = template_result
            st.success("Template validated successfully!")
        else:
            st.error(f"Invalid template: {template_result}")
            template_path = None

# Process button
if uploaded_file and st.button("Generate Presentation"):
    if not setup_gemini_api():
        st.error("Please configure the Gemini API key first.")
    else:
        # Create progress bar
        progress_bar = st.progress(0)
        status_text = st.empty()
        
        # Save uploaded file to temp location
        pdf_bytes = uploaded_file.read()
        with tempfile.NamedTemporaryFile(delete=False, suffix='.pdf') as tmp_file:
            tmp_file.write(pdf_bytes)
            pdf_path = tmp_file.name
        
        try:
            # 1. Extract images from PDF
            status_text.text("Step 1/6: Extracting images from PDF...")
            image_paths_by_page = extract_images_from_pdf(pdf_path)
            progress_bar.progress(1/6)
            
            # 2. Extract text from PDF
            status_text.text("Step 2/6: Extracting text from PDF...")
            text_by_page = extract_text_from_pdf(pdf_path)
            progress_bar.progress(2/6)
            
            # 3. Generate image captions
            status_text.text("Step 3/6: Generating image captions...")
            captions_data = generate_image_captions(image_paths_by_page, text_by_page)
            progress_bar.progress(3/6)
            
            # 4. Optimize prompt based on presentation time
            status_text.text("Step 4/6: Optimizing presentation structure...")
            sample_text = "\n".join([text for page_num, text in sorted(text_by_page.items())][:3])
            optimized_prompt = optimize_prompt_for_presentation_time(user_prompt, sample_text, presentation_time)
            progress_bar.progress(4/6)
            
            # 5. Generate presentation content
            status_text.text("Step 5/6: Generating presentation content...")
            presentation_content = generate_presentation_content(pdf_path, captions_data, optimized_prompt, template_path)
            progress_bar.progress(5/6)
            
            # 6. Create PowerPoint
            status_text.text("Step 6/6: Creating PowerPoint presentation...")
            output_path = create_powerpoint(presentation_content, template_path)
            progress_bar.progress(6/6)
            
            # Complete
            status_text.text("Presentation generation complete!")
            
            # Provide download link
            with open(output_path, "rb") as file:
                st.download_button(
                    label="Download Presentation",
                    data=file,
                    file_name=os.path.basename(output_path),
                    mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
                )
            
            # Display some statistics
            st.success(f"Successfully generated presentation with {len(presentation_content.get('slides', []))} slides and {sum(len(page_images) for page_images in image_paths_by_page.values())} images.")
            
        except Exception as e:
            st.error(f"Error generating presentation: {str(e)}")
        
        finally:
            # Clean up
            try:
                os.unlink(pdf_path)
            except:
                pass
