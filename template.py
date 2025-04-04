import sys
import os
import time
import pandas as pd
import comtypes.client
from comtypes.gen import PowerPoint as PPConst  

if getattr(sys, 'frozen', False): 
    base_path = os.path.dirname(sys.executable)
else: 
    base_path = os.path.dirname(os.path.abspath(__file__))  

print("Running from base path:", base_path)

def rgb_to_ole(red, green, blue):
    """Convert RGB to OLE color format used by PowerPoint."""
    return red + (green * 256) + (blue * 256 * 256)

def duplicate_slide(presentation, slide_index):
    """Duplicate a slide in a PowerPoint presentation."""
    slides = presentation.Slides
    total_slides = slides.Count
    if slide_index < 1 or slide_index > total_slides:
        raise ValueError(f"Invalid slide index: {slide_index}. Must be between 1 and {total_slides}")
    return slides(slide_index).Duplicate().Item(1)

def add_name_to_first_slide(presentation, user_name):
    """Add user's name to the first slide."""
    first_slide = presentation.Slides(1)  
    left_position = first_slide.Master.Width - 450  
    top_position = first_slide.Master.Height - 100  
    text_box = first_slide.Shapes.AddTextbox(1, left_position, top_position, 250, 50)
    text_frame = text_box.TextFrame.TextRange
    text_frame.Text = user_name
    text_frame.Font.Size = 24
    text_frame.Font.Name = "Century Gothic"
    text_frame.Font.Color.RGB = rgb_to_ole(0, 0, 0) 
    text_frame.Font.Bold = True
    text_frame.ParagraphFormat.Alignment = PPConst.ppAlignCenter

def create_anniversary_slides(excel_path, template_path, output_path):
    """Generate a PowerPoint presentation with anniversary wishes."""
    if not os.path.exists(template_path):
        print(f" Error: '{template_path}' not found. Run modify_pptx() first.")
        return
    user_name = input("Enter your name: ").strip() 
    df = pd.read_excel(excel_path).rename(columns=lambda x: x.strip())  
    required_columns = ['Name', 'Wishes']
    if not all(col in df.columns for col in required_columns):
        print(f" Error: Missing required columns in Excel file: {df.columns}")
        return
    try:
        powerpoint = comtypes.client.CreateObject("PowerPoint.Application")
        powerpoint.Visible = 1  
        presentation = powerpoint.Presentations.Open(os.path.abspath(template_path))
        add_name_to_first_slide(presentation, user_name)
        if presentation.Slides.Count < 2:
            raise ValueError("Template must have at least 2 slides (cover + message template).")
        
        # Process first slide (cover)
        first_slide = presentation.Slides(1)
        
        # Process message slides
        message_template_index = 2  
        max_messages_per_slide = 2  
        messages_on_slide = 0  
        current_slide = None  
        text_positions_two = [(100, 100), (100, 250)]
        text_positions_one = [(presentation.PageSetup.SlideWidth // 2 - 250, 
                      presentation.PageSetup.SlideHeight // 3)]

        for index, row in df.iterrows():
            name = " ".join(word.capitalize() for word in str(row['Name']).strip().split() if word)
            message = str(row['Wishes']).strip()

            if not message or message.lower() == 'nan':
                continue  
            is_long_message = len(message) > 150
            if is_long_message or messages_on_slide >= max_messages_per_slide or current_slide is None:
                current_slide = duplicate_slide(presentation, message_template_index)
                messages_on_slide = 0  
            text_positions = text_positions_one if is_long_message else text_positions_two
            text_position = text_positions[messages_on_slide]  
            text_box = current_slide.Shapes.AddTextbox(1, *text_position, 500, 100)
            text_frame = text_box.TextFrame.TextRange
            text_frame.Text = message
            text_frame.Font.Size = 18
            text_frame.Font.Name = "Century Gothic"
            text_frame.Font.Bold = True
            text_frame.Font.Color.RGB = rgb_to_ole(0, 0, 0)
            text_frame.ParagraphFormat.Alignment = PPConst.ppAlignJustify
            message_height = text_box.TextFrame.TextRange.BoundHeight  
            signature_left = text_position[0]  
            signature_top = text_position[1] + message_height + 10  
            signature_box = current_slide.Shapes.AddTextbox(1, signature_left, signature_top, 500, 30)  
            signature_frame = signature_box.TextFrame.TextRange
            signature_frame.Text = f"- {', '.join(name.splitlines())}"  
            signature_frame.Font.Size = 18
            signature_frame.Font.Name = "Century Gothic"
            signature_frame.Font.Color.RGB = rgb_to_ole(255, 0, 0) 
            signature_frame.Font.Italic = True
            signature_frame.ParagraphFormat.Alignment = PPConst.ppAlignRight  
            signature_box.TextFrame.WordWrap = False  
            messages_on_slide += 1 if not is_long_message else max_messages_per_slide  

        # Delete the template slide if it still exists
        if presentation.Slides.Count > 2:
            presentation.Slides(2).Delete()  

        # Add thank you slide (will be slide 3)
        add_thank_you_slide(presentation)

        presentation.SaveAs(os.path.abspath(output_path))
        presentation.Close()
        powerpoint.Quit()
        print(f"Presentation created successfully: {output_path}")

    except Exception as e:
        print(f"Error in create_anniversary_slides: {e}")

def add_thank_you_slide(presentation):
    """Add 'Thank You' text to the last slide only."""
    thank_you_slide = presentation.Slides(1).Duplicate().Item(1)
    thank_you_slide.MoveTo(presentation.Slides.Count)

    for shape in list(thank_you_slide.Shapes):  
        if not (shape.Type == 13 and shape.Left < 100 and shape.Top < 100):
            shape.Delete()
    text_box = thank_you_slide.Shapes.AddTextbox(
        1, 
        presentation.PageSetup.SlideWidth/2 - 150,  
        presentation.PageSetup.SlideHeight/2 - 40,  
        300, 80
    )
    text_frame = text_box.TextFrame.TextRange
    text_frame.Text = "Thank You"
    text_frame.Font.Size = 65
    text_frame.Font.Name = "Century Gothic"
    text_frame.Font.Color.RGB = rgb_to_ole(255, 0, 0)  
    text_frame.Font.Italic = True
    text_frame.ParagraphFormat.Alignment = PPConst.ppAlignCenter
    text_box.TextFrame.WordWrap = False

def modify_pptx(input_path, output_path):
    """Modify a PowerPoint template while preserving the Entrans logo on all slides"""
    input_path = os.path.abspath(input_path)
    output_path = os.path.abspath(output_path)
    if not os.path.exists(input_path):
        print(f"Error: Input PowerPoint file not found: {input_path}")
        return
    if os.path.exists(output_path):
        try:
            os.remove(output_path)  
            print(f"Closed existing file: {output_path}")
        except PermissionError:
            print(f"Cannot modify {output_path}, it's open. Close it and try again.")
            return
    try:
        powerpoint = comtypes.client.CreateObject("PowerPoint.Application")
        powerpoint.Visible = 1  
        presentation = powerpoint.Presentations.Open(input_path)
        selected_anniversary = int(input("Enter the Work Anniversary (1, 2, 3, or 4): "))
        anniversary_slides = {1: 1, 2: 2, 3: 3, 4: 4}  
        if selected_anniversary not in anniversary_slides:
            print("Invalid anniversary number!")
            return

        # Find and store the logo from the first slide
        logo_shape = None
        first_slide = presentation.Slides(1)
        for shape in first_slide.Shapes:
            if shape.Type == 13 and shape.Left < 100 and shape.Top < 100:  # Assuming logo is in top-left
                logo_shape = shape
                break

        # Delete unwanted anniversary slides
        for i in reversed(range(1, presentation.Slides.Count + 1)):
            if i != selected_anniversary:
                presentation.Slides(i).Delete()

        # Process first slide - keep both logo and anniversary image
        first_slide = presentation.Slides(1)
        shapes_to_delete = []
        anniversary_image = None
        
        for shape in first_slide.Shapes:
            if shape.HasTextFrame and ("ANNIVERSARY" in shape.TextFrame.TextRange.Text.upper() or 
                                    "Employee Name" in shape.TextFrame.TextRange.Text):
                shapes_to_delete.append(shape)
            elif shape.Type == 13:
                # Keep both logo (top-left) and anniversary image (assuming centered)
                if not (shape.Left < 100 and shape.Top < 100):
                    anniversary_image = shape

        for shape in shapes_to_delete:
            shape.Delete()

        # Create message slide (will be slide 2)
        message_slide = first_slide.Duplicate().Item(1)
        
        # Clean up message slide (keep only logo)
        for shape in list(message_slide.Shapes):
            if not (shape.Type == 13 and shape.Left < 100 and shape.Top < 100):
                shape.Delete()

        temp_output = output_path.replace(".pptx", "_temp.pptx")
        presentation.SaveAs(temp_output)
        presentation.Close()
        time.sleep(1)
        os.replace(temp_output, output_path)
        print(f"Modified PowerPoint saved as: {output_path}")
    except Exception as e:
        print(f"Error in modify_pptx: {e}")
    finally:
        powerpoint.Quit()

if __name__ == "__main__":
    input_pptx = os.path.join(base_path, "WorkAnniversaryLogo.pptx")
    output_pptx = os.path.join(base_path, "Modified_Work_Anniversary.pptx")
    excel_file = os.path.join(base_path, "anniversary.xlsx")
    final_pptx = os.path.join(base_path, "Final_Anniversary_Presentation.pptx")  
    
    if not os.path.exists(input_pptx):
        print(f" Error: Input PowerPoint file not found: {input_pptx}")
    else:
        modify_pptx(input_pptx, output_pptx)  

    time.sleep(2)  
    if os.path.exists(output_pptx):
        create_anniversary_slides(excel_file, output_pptx, final_pptx)
        print(f" Final presentation saved at: {final_pptx}")
    else:
        print(" Error: Modified PowerPoint file was not created. Exiting.")