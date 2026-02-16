import win32clipboard
import win32con
from PIL import Image
import ctypes
from ctypes import wintypes
import tempfile
import os
import sys
import subprocess

def save_emf_as_png(output_path):
    """Save EMF from clipboard as PNG file"""
    temp_emf_path = None
    
    # Define handle type based on platform
    if sys.maxsize > 2**32:  # 64-bit
        HANDLE = ctypes.c_void_p
    else:  # 32-bit
        HANDLE = ctypes.c_uint32
    
    try:
        # Open clipboard
        win32clipboard.OpenClipboard()
        
        # Check if EMF format is available
        if not win32clipboard.IsClipboardFormatAvailable(win32con.CF_ENHMETAFILE):
            print("CF_ENHMETAFILE not available")
            win32clipboard.CloseClipboard()
            return False
        
        print("CF_ENHMETAFILE is available")
        
        # Get handle via ctypes
        GetClipboardData = ctypes.windll.user32.GetClipboardData
        GetClipboardData.argtypes = [wintypes.UINT]
        GetClipboardData.restype = HANDLE
        
        emf_handle = GetClipboardData(win32con.CF_ENHMETAFILE)
        print(f"EMF handle: {emf_handle}")
        
        # Create temp file for EMF
        temp_emf = tempfile.NamedTemporaryFile(suffix='.emf', delete=False)
        temp_emf_path = temp_emf.name
        temp_emf.close()
        
        # Copy EMF to file using CopyEnhMetaFileW
        gdi32 = ctypes.windll.gdi32
        CopyEnhMetaFileW = gdi32.CopyEnhMetaFileW
        CopyEnhMetaFileW.argtypes = [HANDLE, wintypes.LPCWSTR]
        CopyEnhMetaFileW.restype = HANDLE
        
        new_handle = CopyEnhMetaFileW(emf_handle, temp_emf_path)
        
        # Close clipboard
        win32clipboard.CloseClipboard()
        
        if not new_handle:
            print("Failed to copy EMF to file")
            return False
        
        print(f"EMF saved to: {temp_emf_path}")
        
        # Delete the handle (we have the file now)
        DeleteEnhMetaFile = gdi32.DeleteEnhMetaFile
        DeleteEnhMetaFile.argtypes = [HANDLE]
        DeleteEnhMetaFile.restype = wintypes.BOOL
        DeleteEnhMetaFile(new_handle)
        
        # Now convert EMF to PNG using PIL
        # PIL can open EMF files on Windows
        try:
            print("Opening EMF with PIL...")
            img = Image.open(temp_emf_path)
            print(f"EMF opened successfully. Size: {img.size}, Mode: {img.mode}")
            
            # Convert to RGB if necessary
            if img.mode != 'RGB':
                img = img.convert('RGB')
            
            # Save as PNG
            img.save(output_path, 'PNG')
            print(f"Successfully saved to {output_path}")
            
            return True
            
        except Exception as pil_error:
            print(f"PIL failed to open EMF: {pil_error}")
            print("Trying alternative method with ImageMagick...")
            
            # Try using ImageMagick if available
            try:
                result = subprocess.run(
                    ['magick', 'convert', temp_emf_path, output_path],
                    capture_output=True,
                    text=True,
                    timeout=30
                )
                if result.returncode == 0:
                    print(f"Successfully converted with ImageMagick to {output_path}")
                    return True
                else:
                    print(f"ImageMagick failed: {result.stderr}")
            except FileNotFoundError:
                print("ImageMagick not found")
            except Exception as img_error:
                print(f"ImageMagick error: {img_error}")
            
            print("\nThe EMF file has been saved. You can:")
            print(f"1. Open it directly: {temp_emf_path}")
            print("2. Convert it manually using another tool")
            print("3. Install ImageMagick and try again")
            
            return False
            
    except Exception as e:
        print(f"Error: {e}")
        import traceback
        traceback.print_exc()
        try:
            win32clipboard.CloseClipboard()
        except:
            pass
        return False
    
    finally:
        # Clean up temp file if conversion was successful
        if temp_emf_path and os.path.exists(temp_emf_path):
            try:
                # Only delete if we successfully created the PNG
                if os.path.exists(output_path):
                    os.unlink(temp_emf_path)
                    print(f"Cleaned up temp file")
            except:
                pass

# Usage
if __name__ == "__main__":
    result = save_emf_as_png("output.png")
    if result:
        print("\n✓ Success!")
    else:
        print("\n✗ Failed")