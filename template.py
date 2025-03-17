import os
import pandas as pd
import comtypes.client
from comtypes.gen import PowerPoint as PPConst  

def duplicate_slide(presentation, slide_index):
    """Duplicate a slide in a PowerPoint presentation, adjusting for zero-based index."""
    slides = presentation.Slides
    total_slides = slides.Count

    if slide_index < 0 or slide_index >= total_slides:
        raise ValueError(f"Invalid slide index: {slide_index}. Must be between 0 and {total_slides - 1}")

    duplicated_slide = slides(slide_index + 1).Duplicate().Item(1) 
    return duplicated_slide

def rgb_to_ole(red, green, blue):
    """Convert RGB to OLE color format used by PowerPoint."""
    return red + (green * 256) + (blue * 256 * 256)

def add_name_to_first_slide(presentation, user_name):
    """Add the user-defined name to the bottom-right corner of the first slide."""
    first_slide = presentation.Slides(1) 

    # Define position for bottom-right corner
    left_position = first_slide.Master.Width - 450  
    top_position = first_slide.Master.Height - 100  

    # Add the text box
    text_box = first_slide.Shapes.AddTextbox(1, left_position, top_position, 250, 50)
    text_frame = text_box.TextFrame.TextRange

    # Set text properties
    text_frame.Text = user_name
    text_frame.Font.Size = 20
    text_frame.Font.Name = "Arial"
    text_frame.Font.Color.RGB = rgb_to_ole(0, 0, 0) 
    text_frame.Font.Bold = True
    text_frame.ParagraphFormat.Alignment = PPConst.ppAlignRight  

def create_anniversary_slides(excel_path, template_path, output_path):
    """Generate a PowerPoint presentation with anniversary wishes."""
    
    user_name = input("Enter your name: ").strip() 
    df = pd.read_excel(excel_path)
    df = df.rename(columns=lambda x: x.strip())  

    # Validate column names
    required_columns = ['Name', 'Wishes']
    for col in required_columns:
        if col not in df.columns:
            print(f"Error: Missing column '{col}' in Excel file.")
            print("Available columns:", df.columns)
            return
    powerpoint = comtypes.client.CreateObject("PowerPoint.Application")
    powerpoint.Visible = 1  

    # Open template
    presentation = powerpoint.Presentations.Open(os.path.abspath(template_path))
    add_name_to_first_slide(presentation, user_name)
    if presentation.Slides.Count < 2:
        raise ValueError("Template must have at least 2 slides (cover + message template).")

    message_template_index = 1  

    max_messages_per_slide = 2
    messages_on_slide = 0
    current_slide = None

    # Define text positions
    text_box_positions_two = [(100, 100), (100, 250)] 
    text_box_positions_one = [(100, 250)]  
    text_box_positions_single_lower = [(100, 350)]  

    for index, row in df.iterrows():
        name = str(row['Name']).strip()
        message = str(row['Wishes']).strip()

        if not message or message.lower() == 'nan':
            continue 

        is_long_message = len(message) > 150

        # Create a new slide if needed
        if is_long_message or messages_on_slide >= max_messages_per_slide or current_slide is None:
            current_slide = duplicate_slide(presentation, message_template_index)
            messages_on_slide = 0  
        if messages_on_slide == 0 and index == len(df) - 1:
            text_box_positions = text_box_positions_single_lower  
        else:
            text_box_positions = text_box_positions_two 

        # Add text box
        text_box = current_slide.Shapes.AddTextbox(1, *text_box_positions[messages_on_slide], 500, 100)
        text_frame = text_box.TextFrame.TextRange
        text_frame.Text = message
        text_frame.Font.Size = 18
        text_frame.Font.Name = "Arial"
        text_frame.Font.Color.RGB = rgb_to_ole(0, 0, 0) 
        name_text = f"\n\n-{name}"
        text_frame.Text += name_text
        name_start = len(text_frame.Text) - len(name_text) + 1
        name_length = len(name_text)
        for i in range(name_length):
            text_frame.Characters(name_start + i, 1).Font.Color.RGB = rgb_to_ole(255, 0, 0) 
            text_frame.Characters(name_start + i, 1).Font.Bold = True  
        text_frame.ParagraphFormat.Alignment = PPConst.ppAlignCenter  
        messages_on_slide += 1
        if is_long_message:
            messages_on_slide = max_messages_per_slide 

    # Delete the second slide after placing all messages (zero-based index 1)
    if presentation.Slides.Count > 1:
        presentation.Slides(2).Delete()  
    presentation.SaveAs(os.path.abspath(output_path))
    presentation.Close()
    powerpoint.Quit()
    print("Presentation created successfully!")
create_anniversary_slides(
    "anniversary.xlsx",
    "template.pptx",
    "Anniversary.pptx"
)
