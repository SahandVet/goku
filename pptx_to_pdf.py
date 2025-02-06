import os
import argparse
from win32com import client

def convert_pptx_to_pdf(input_path, output_path):
    powerpoint = client.Dispatch("PowerPoint.Application")
    try:
        presentation = powerpoint.Presentations.Open(input_path)
        presentation.SaveAs(output_path, 32)  # 32 is the PDF format
        presentation.Close()
    except Exception as e:
        print(f"Error converting {input_path}: {str(e)}")
    finally:
        powerpoint.Quit()

def main():
    parser = argparse.ArgumentParser(description='Convert PPTX files to PDF')
    parser.add_argument('-i', '--input', required=True, help='Input directory containing PPTX files')
    parser.add_argument('-o', '--output', required=True, help='Output directory for PDF files')
    
    args = parser.parse_args()
    
    if not os.path.exists(args.output):
        os.makedirs(args.output)
    
    for filename in os.listdir(args.input):
        if filename.endswith(".pptx"):
            input_path = os.path.join(args.input, filename)
            output_filename = os.path.splitext(filename)[0] + ".pdf"
            output_path = os.path.join(args.output, output_filename)
            
            print(f"Converting {filename}...")
            convert_pptx_to_pdf(input_path, output_path)
    
    print("Conversion complete!")

if __name__ == "__main__":
    main()
