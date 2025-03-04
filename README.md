# StreamSlide

StreamSlide is an AI-powered tool that transforms PDF documents into professional PowerPoint presentations with minimal effort. Leveraging Google's Gemini API, it automatically extracts content, captions images, and generates structured presentations.

![StreamSlide Demo](https://placeholder-for-demo-image.png)

## Features

- **PDF Processing**: Extracts both text and images from PDF documents
- **AI Image Captioning**: Uses Gemini Vision to generate relevant captions for images
- **Smart Content Generation**: Creates coherent presentation content based on document context
- **Custom PowerPoint Creation**: Builds professional slides with appropriate text and image placement
- **Presentation Templates**: Choose from multiple presentation styles
- **User-Friendly Interface**: Simple Streamlit interface with progress tracking
- **API Key Management**: Use environment variables or enter your API key directly

## Installation

### Prerequisites

- Python 3.8+
- Google Gemini API key ([Get one here](https://ai.google.dev/))

### Setup

1. Clone the repository:

   ```bash
   git clone https://github.com/yourusername/StreamSlide.git
   cd StreamSlide
   ```
2. Create a virtual environment:

   ```bash
   python -m venv venv
   ```
3. Activate the virtual environment:

   - Windows:
     ```bash
     venv\Scripts\activate
     ```
   - macOS/Linux:
     ```bash
     source venv/bin/activate
     ```
4. Install dependencies:

   ```bash
   pip install streamlit google-generativeai langchain langchain-community pypdf python-pptx Pillow python-dotenv PyMuPDF requests
   ```
5. (Optional) Create a .env file with your Gemini API key:

   ```bash
   echo "GEMINI_API_KEY=your_api_key_here" > .env
   ```

## Usage

1. Start the application:

   ```bash
   streamlit run app.py
   ```
2. Access the application in your web browser (typically at http://localhost:8501)
3. Enter your Gemini API key if not already set in the .env file
4. Upload a PDF document
5. Describe what you want in the presentation
6. Select a presentation template style
7. Click "Generate Presentation" and watch the magic happen
8. Download your professionally created PowerPoint presentation

## How It Works

StreamSlide processes your PDF in five key steps:

1. **PDF Processing**: Extracts text and images from the uploaded PDF
2. **Image Analysis**: Uses Gemini Vision to understand and caption each image
3. **Content Generation**: Creates presentation structure and content based on the document
4. **Slide Creation**: Builds PowerPoint slides with appropriate layouts
5. **Presentation Assembly**: Combines everything into a downloadable PowerPoint file

## Project Structure

StreamSlide/

│

├── app.py # Main application file

├── .env # Environment variables (optional)

├── README.md # Project documentation

├── requirements.txt # Python dependencies

│

├── temp_files/ # Temporary processing files

│ └── images/ # Extracted images

│

└── output/ # Generated PowerPoint files


## Requirements

- streamlit
- google-generativeai
- langchain
- langchain-community
- python-pptx
- Pillow
- python-dotenv
- PyMuPDF
- requests

## Limitations

- Processing large PDFs may take time and consume API tokens
- Complex images may receive generic captions
- The free tier of Gemini API has rate limits
- Only supports PDF files as input

## Contributing

Contributions are welcome! Please feel free to submit a Pull Request.

1. Fork the repository
2. Create your feature branch (`git checkout -b feature/amazing-feature`)
3. Commit your changes (`git commit -m 'Add some amazing feature'`)
4. Push to the branch (`git push origin feature/amazing-feature`)
5. Open a Pull Request

## License

This project is licensed under the MIT License - see the LICENSE file for details.

## Acknowledgments

- Google Gemini for AI capabilities
- Streamlit for the interactive web interface
- PyMuPDF for PDF processing
- python-pptx for PowerPoint generation
