import streamlit as st
from pptx import Presentation

def extract_text_from_pptx(uploaded_file):
    """Extract text from uploaded PPTX file"""
    try:
        # Create a Presentation object from the uploaded file
        ppt = Presentation(uploaded_file)
        
        extracted_text = []
        slide_count = 0
        
        for slide in ppt.slides:
            slide_count += 1
            slide_text = f"\n--- Slide {slide_count} ---\n"
            
            for shape in slide.shapes:
                if hasattr(shape, "text") and shape.text.strip():
                    slide_text += shape.text + "\n"
            
            extracted_text.append(slide_text)
        
        return "\n".join(extracted_text), slide_count
    
    except Exception as e:
        return f"Error processing file: {str(e)}", 0

# Set up the web app interface
st.title("🎯 PPTX to Text Converter")
st.markdown("**Transform your PowerPoint presentations into clean text format**")

# File upload section
st.subheader("Upload Your Presentation")
uploaded_file = st.file_uploader(
    "Choose a PPTX file", 
    type=['pptx'],
    help="Select a PowerPoint file (.pptx) to extract text from"
)

if uploaded_file is not None:
    # Display file info
    st.success(f"File uploaded: {uploaded_file.name}")
    st.info(f"File size: {uploaded_file.size:,} bytes")
    
    # Process the file when user clicks the button
    if st.button("🔄 Extract Text", type="primary"):
        with st.spinner("Processing your presentation..."):
            extracted_text, slide_count = extract_text_from_pptx(uploaded_file)
        
        if slide_count > 0:
            st.success(f"✅ Successfully processed {slide_count} slides!")
            
            # Display the extracted text
            st.subheader("Extracted Text")
            st.text_area(
                "Text Content", 
                extracted_text, 
                height=400,
                help="You can copy this text using Ctrl+C"
            )
            
            # Download option
            st.download_button(
                label="📥 Download as Text File",
                data=extracted_text,
                file_name=f"{uploaded_file.name.replace('.pptx', '')}_extracted.txt",
                mime="text/plain"
            )
        else:
            st.error("❌ Could not extract text from the file. Please check if it's a valid PPTX file.")

# Instructions section
with st.expander("ℹ️ How to Use"):
    st.markdown("""
    1. **Upload**: Click 'Browse files' and select your .pptx file
    2. **Process**: Click 'Extract Text' to convert your presentation
    3. **Review**: Read the extracted text in the text area below
    4. **Download**: Save the text as a .txt file to your computer
    
    **Note**: This tool extracts text from text boxes, titles, and content in your slides.
    """)

# Footer
st.markdown("---")
st.markdown("*Built with Streamlit • Simple, Fast, Effective*")