import os
import shutil
import win32com.client

def convert_ppt_to_png(ppt_path, output_folder):
    if not os.path.exists(ppt_path):
        raise FileNotFoundError(f"PPT file not found: {ppt_path}")

    # Clear old images if folder exists
    if os.path.exists(output_folder):
        shutil.rmtree(output_folder)
    os.makedirs(output_folder, exist_ok=True)

    # Start PowerPoint
    powerpoint = win32com.client.Dispatch("PowerPoint.Application")
    powerpoint.Visible = 1

    # Open in read-only mode (avoid conflicts)
    presentation = powerpoint.Presentations.Open(ppt_path, WithWindow=False)

    # Export slides (18 = PNG)
    presentation.SaveAs(output_folder, 18)

    # Close PowerPoint
    presentation.Close()
    powerpoint.Quit()

    print(f"✅ Slides saved as PNG in: {output_folder}")

# Example Usage
ppt_file = r"C:\Users\harin\OneDrive\Desktop\review2.pptx"
output_dir = r"C:\Users\harin\OneDrive\Desktop\PROJECTS-G\Hand Gesture Based Virtual presentation system\OutputFolder"

convert_ppt_to_png(ppt_file, output_dir)

