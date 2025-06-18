import os
import streamlit as st
from PIL import Image
from io import BytesIO
import zipfile
import tempfile

# Handle docx import with fallback
try:
    from docx import Document
    DOCX_AVAILABLE = True
except ImportError as e:
    st.error(f"python-docx import failed: {e}")
    st.info("Install with: pip install python-docx")
    DOCX_AVAILABLE = False
except Exception as e:
    st.error(f"DOCX module conflict detected: {e}")
    st.info("Try: pip uninstall docx && pip install python-docx")
    DOCX_AVAILABLE = False

def is_likely_graph(image, width, height, filename=""):
    """Enhanced graph detection optimized for technical reports"""
    # Minimum size check - be more lenient for technical diagrams
    if width < 100 or height < 80:
        return False
    
    # Don't exclude smaller technical charts/diagrams
    if width * height < 8000:  # Very small icons only
        return False
    
    aspect_ratio = width / height
    
    # More flexible aspect ratios for technical content
    if not (0.3 <= aspect_ratio <= 8.0):
        return False
    
    try:
        # Multi-criteria analysis for technical content
        gray = image.convert('L')
        pixels = list(gray.getdata())
        total_pixels = len(pixels)
        
        if total_pixels == 0:
            return False
        
        # Analyze brightness distribution
        avg_brightness = sum(pixels) / total_pixels
        
        # Technical charts often have varied brightness patterns
        brightness_bins = [0, 0, 0, 0]  # dark, medium-dark, medium-light, light
        for pixel in pixels:
            if pixel < 64:
                brightness_bins[0] += 1
            elif pixel < 128:
                brightness_bins[1] += 1
            elif pixel < 192:
                brightness_bins[2] += 1
            else:
                brightness_bins[3] += 1
        
        # Normalize to percentages
        brightness_dist = [b/total_pixels for b in brightness_bins]
        
        # Charts typically have varied brightness (not all one shade)
        max_single_brightness = max(brightness_dist)
        if max_single_brightness < 0.85:  # Not dominated by single brightness
            return True
        
        # Check for edge content (typical in charts with axes/borders)
        edge_pixels = []
        sample_size = min(width, height) // 10
        
        # Sample edges
        for i in range(0, width, max(1, width//20)):
            if i < width:
                edge_pixels.extend([gray.getpixel((i, 0)), gray.getpixel((i, height-1))])
        for i in range(0, height, max(1, height//20)):
            if i < height:
                edge_pixels.extend([gray.getpixel((0, i)), gray.getpixel((width-1, i))])
        
        if edge_pixels:
            edge_variance = sum((p - avg_brightness)**2 for p in edge_pixels) / len(edge_pixels)
            if edge_variance > 500:  # Varied edges suggest chart content
                return True
        
        # Check for high contrast regions (typical in technical diagrams)
        sorted_pixels = sorted(pixels)
        if len(sorted_pixels) >= 4:
            q1 = sorted_pixels[len(sorted_pixels)//4]
            q3 = sorted_pixels[3*len(sorted_pixels)//4]
            if q3 - q1 > 80:  # Good contrast range
                return True
        
        # Check for moderate complexity (not solid color, not pure noise)
        unique_colors = len(set(pixels[:1000]))  # Sample for performance
        color_complexity = unique_colors / min(1000, total_pixels)
        if 0.1 <= color_complexity <= 0.8:  # Moderate complexity
            return True
            
    except Exception as e:
        # If analysis fails, use size-based fallback
        pass
    
    # Size-based fallback for larger images
    return width >= 200 and height >= 120

def extract_images_from_docx(docx_path, output_folder):
    """Extract images using multiple methods for better compatibility"""
    extracted_count = 0
    
    # Always try ZIP method first (more reliable)
    extracted_count = extract_via_zip(docx_path, output_folder, 0)
    
    # Try python-docx method if available and ZIP method found nothing
    if extracted_count == 0 and DOCX_AVAILABLE:
        try:
            st.info("🔄 Trying python-docx extraction method...")
            doc = Document(docx_path)
            
            if hasattr(doc.part, 'rels'):
                for rel_id, rel in doc.part.rels.items():
                    if "image" in rel.target_ref.lower():
                        try:
                            image_part = rel.target_part
                            image_data = image_part.blob
                            
                            img = Image.open(BytesIO(image_data))
                            width, height = img.size
                            
                            if is_likely_graph(img, width, height):
                                extracted_count += 1
                                
                                # Determine file extension
                                content_type = getattr(image_part, 'content_type', 'image/png')
                                ext = content_type.split('/')[-1].lower()
                                if ext == 'jpeg': ext = 'jpg'
                                if ext not in ['png', 'jpg', 'gif', 'bmp']: ext = 'png'
                                
                                filename = f"graph_{extracted_count:03d}.{ext}"
                                filepath = os.path.join(output_folder, filename)
                                
                                # Convert mode if needed
                                if ext in ['jpg', 'jpeg'] and img.mode in ['RGBA', 'P']:
                                    img = img.convert('RGB')
                                
                                img.save(filepath, quality=95, optimize=True)
                                st.success(f"✅ Extracted: {filename} ({width}×{height})")
                            else:
                                st.info(f"⚪ Skipped small/non-graph image ({width}×{height})")
                                
                        except Exception as e:
                            st.warning(f"⚠️ Error processing image: {str(e)}")
                            continue
                            
        except Exception as e:
            st.error(f"❌ python-docx method failed: {str(e)}")
    
    return extracted_count

def extract_via_zip(docx_path, output_folder, start_count):
    """Enhanced ZIP extraction with better error handling and filtering"""
    count = 0
    processed_images = set()  # Avoid duplicates
    
    try:
        with zipfile.ZipFile(docx_path, 'r') as docx_zip:
            # Get all image files
            image_files = [f for f in docx_zip.infolist() 
                          if f.filename.startswith('word/media/') and 
                          not f.filename.endswith('/')]
            
            st.info(f"🔍 Found {len(image_files)} images in document")
            
            for file_info in image_files:
                try:
                    # Skip if already processed (handle duplicates)
                    file_key = f"{file_info.filename}_{file_info.file_size}"
                    if file_key in processed_images:
                        continue
                    processed_images.add(file_key)
                    
                    with docx_zip.open(file_info) as img_file:
                        img_data = img_file.read()
                        
                        # Validate image data
                        if len(img_data) < 100:  # Skip tiny files
                            st.info(f"⚪ Skipped tiny file: {file_info.filename}")
                            continue
                            
                        try:
                            img = Image.open(BytesIO(img_data))
                            width, height = img.size
                            
                            # More detailed logging
                            original_name = os.path.basename(file_info.filename)
                            st.info(f"📊 Analyzing: {original_name} ({width}×{height})")
                            
                            if is_likely_graph(img, width, height, file_info.filename):
                                count += 1
                                
                                # Better extension handling
                                ext = original_name.split('.')[-1].lower() if '.' in original_name else 'png'
                                if ext not in ['png', 'jpg', 'jpeg', 'gif', 'bmp', 'tiff']:
                                    # Try to detect format from image
                                    try:
                                        format_map = {'JPEG': 'jpg', 'PNG': 'png', 'GIF': 'gif', 
                                                    'BMP': 'bmp', 'TIFF': 'tiff'}
                                        ext = format_map.get(img.format, 'png')
                                    except:
                                        ext = 'png'
                                
                                filename = f"chart_{start_count + count:03d}_{original_name.split('.')[0]}.{ext}"
                                filepath = os.path.join(output_folder, filename)
                                
                                # Handle image mode conversion
                                try:
                                    if ext in ['jpg', 'jpeg']:
                                        if img.mode in ['RGBA', 'P', 'LA']:
                                            # Create white background for transparency
                                            bg = Image.new('RGB', img.size, (255, 255, 255))
                                            if img.mode == 'P':
                                                img = img.convert('RGBA')
                                            bg.paste(img, mask=img.split()[-1] if img.mode in ['RGBA', 'LA'] else None)
                                            img = bg
                                        elif img.mode != 'RGB':
                                            img = img.convert('RGB')
                                    elif ext == 'png' and img.mode not in ['RGB', 'RGBA', 'P', 'L']:
                                        img = img.convert('RGBA')
                                    
                                    # Save with optimization
                                    save_kwargs = {'optimize': True}
                                    if ext in ['jpg', 'jpeg']:
                                        save_kwargs['quality'] = 95
                                        save_kwargs['dpi'] = (300, 300)
                                    elif ext == 'png':
                                        save_kwargs['dpi'] = (300, 300)
                                    
                                    img.save(filepath, **save_kwargs)
                                    
                                    # Verify file was saved
                                    if os.path.exists(filepath) and os.path.getsize(filepath) > 0:
                                        st.success(f"✅ Extracted: {filename} ({width}×{height})")
                                    else:
                                        st.error(f"❌ Failed to save: {filename}")
                                        count -= 1
                                        
                                except Exception as save_error:
                                    st.warning(f"⚠️ Save error for {filename}: {str(save_error)}")
                                    count -= 1
                            else:
                                reason = "size" if (width < 200 or height < 120) else "content analysis"
                                st.info(f"⚪ Skipped ({reason}): {original_name} ({width}×{height})")
                                
                        except Exception as img_error:
                            st.warning(f"⚠️ Image processing error for {file_info.filename}: {str(img_error)}")
                            continue
                            
                except Exception as file_error:
                    st.warning(f"⚠️ File access error for {file_info.filename}: {str(file_error)}")
                    continue
                    
    except Exception as e:
        st.error(f"❌ ZIP extraction failed: {str(e)}")
        
    return count

def create_download_zip(output_folder):
    """Create a ZIP file of extracted images for download"""
    if not os.path.exists(output_folder) or not os.listdir(output_folder):
        return None
    
    zip_path = f"{output_folder}.zip"
    
    with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED) as zipf:
        for filename in os.listdir(output_folder):
            if filename.lower().endswith(('.png', '.jpg', '.jpeg', '.gif', '.bmp')):
                file_path = os.path.join(output_folder, filename)
                zipf.write(file_path, filename)
    
    return zip_path

def main():
    st.set_page_config(
        page_title="DOCX Graph Extractor",
        page_icon="📈",
        layout="wide"
    )
    
    st.title("📈 DOCX Graph Extractor")
    st.markdown("""
    **Extract graphs and charts from Word documents with intelligent filtering.**
    
    This tool automatically identifies and extracts images that are likely to be graphs, charts, 
    or diagrams based on size, aspect ratio, and visual characteristics.
    """)
    
    # Sidebar for settings
    with st.sidebar:
        st.header("⚙️ Settings")
        output_folder = st.text_input(
            "Output folder name", 
            value="extracted_graphs",
            help="Name of the folder where images will be saved"
        )
        
        st.markdown("---")
        st.markdown("**Detection Criteria:**")
        st.markdown("• Minimum size: 100×80 px")
        st.markdown("• Flexible aspect ratios") 
        st.markdown("• Content analysis for charts")
        st.markdown("• Edge detection & contrast analysis")
        
        st.markdown("---")
        st.markdown("**Chart Types Detected:**")
        st.markdown("• Line graphs & time series")
        st.markdown("• Bar charts & histograms") 
        st.markdown("• Technical diagrams")
        st.markdown("• Network topology charts")
        st.markdown("• Compliance reports graphs")
    
    # Main interface
    col1, col2 = st.columns([2, 1])
    
    with col1:
        uploaded_file = st.file_uploader(
            "📁 Choose a DOCX file",
            type=["docx"],
            help="Upload a Word document containing graphs or charts"
        )
    
    with col2:
        if uploaded_file:
            st.info(f"**File:** {uploaded_file.name}")
            st.info(f"**Size:** {uploaded_file.size:,} bytes")
    
    if uploaded_file and output_folder:
        if not DOCX_AVAILABLE:
            st.error("❌ Cannot process DOCX files. Please fix the python-docx installation.")
            st.code("pip uninstall docx\npip install python-docx")
            return
            
        # Create output directory
        os.makedirs(output_folder, exist_ok=True)
        
        if st.button("🚀 Extract Graphs", type="primary", use_container_width=True):
            with st.spinner("Processing document..."):
                # Save uploaded file temporarily
                with tempfile.NamedTemporaryFile(delete=False, suffix='.docx') as tmp_file:
                    tmp_file.write(uploaded_file.getbuffer())
                    temp_path = tmp_file.name
                
                try:
                    # Extract images
                    extracted_count = extract_images_from_docx(temp_path, output_folder)
                    
                    if extracted_count > 0:
                        st.success(f"🎉 Successfully extracted {extracted_count} graph(s)!")
                        
                        # Show extracted files
                        st.subheader("📋 Extracted Files:")
                        files = [f for f in os.listdir(output_folder) 
                                if f.lower().endswith(('.png', '.jpg', '.jpeg', '.gif', '.bmp'))]
                        
                        if files:
                            # Display files in columns
                            cols = st.columns(min(3, len(files)))
                            for i, filename in enumerate(files):
                                with cols[i % 3]:
                                    file_path = os.path.join(output_folder, filename)
                                    try:
                                        img = Image.open(file_path)
                                        st.image(img, caption=filename, use_container_width=True)
                                    except Exception:
                                        st.text(f"📄 {filename}")
                            
                            # Create download option
                            zip_path = create_download_zip(output_folder)
                            if zip_path and os.path.exists(zip_path):
                                with open(zip_path, 'rb') as f:
                                    st.download_button(
                                        label="📥 Download All Images (ZIP)",
                                        data=f.read(),
                                        file_name=f"{output_folder}.zip",
                                        mime="application/zip",
                                        use_container_width=True
                                    )
                                os.remove(zip_path)  # Clean up
                    else:
                        st.warning("⚠️ No graphs found in the document. The file might not contain charts or the images might not meet the detection criteria.")
                        
                finally:
                    # Clean up temp file
                    if os.path.exists(temp_path):
                        os.remove(temp_path)

if __name__ == "__main__":
    main()