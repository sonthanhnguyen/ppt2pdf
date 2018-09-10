"""
    A simple Python script to convert all Powerpoint files in a certain directory to
    pdf files in Windows environment.

    Usage: python ppt2pdf.py -p <path_to_source_directory>

    Dependency:
        The script uses comtypes. Install it by 'pip install comtypes'
"""
import os
import sys
import comtypes.client


def ppt_2_pdf(input_ppt_file, output_pdf_file, format_type=32):
    """
    Convert a Powerpoint file to a pdf file

    :param input_ppt_file: input Powerpoint file
    :param output_pdf_file: output pdf file
    :param format_type:
    :return: a pdf file written in the directory
    """
    powerpoint = comtypes.client.CreateObject("Powerpoint.Application")
    powerpoint.Visible = 1

    if output_pdf_file[-3:] != 'pdf':
        output_pdf_file = output_pdf_file + ".pdf"

    ppt_file = powerpoint.Presentations.Open(input_ppt_file)
    ppt_file.SaveAs(output_pdf_file, format_type)
    ppt_file.Close()
    powerpoint.Quit()


def convert_all_ppt(directory):
    """
    Convert all Powerpoint files in the same directory to pdf files

    :param directory: the full path to the directory
    :return: all pdf outputs in the same directory
    """
    try:
        for file in os.listdir(directory):
            _, file_extension = os.path.splitext(file)
            if "ppt" in file_extension:
                input_file = directory + "\\" + file
                output_file = input_file + "_output.pdf"
                ppt_2_pdf(input_file, output_file)
    except FileNotFoundError:
        print("The system cannot file directory \'{0}\'".format(directory))
        exit(2)
    except OSError:
        print("The filename, directory name, or volume label syntax is incorrect: \'{0}\'"
              .format(directory))
        exit(2)


def main(argv):
    """
    Main function
    """
    # No argument
    if len(argv) <= 1:
        print("To get help, use: ppt2pdf.py -h")
        sys.exit()

    parameter = argv[1]

    if parameter == "-h":
        print("Usage: ppt2pdf.py -p <directory_path>")
    # Convert
    elif parameter == "-p":
        if len(argv) < 3:
            print("No path to source directory. To get help, use: ppt2pdf.py -h")
        else:
            # Only take the third argument, ignore the rest
            directory_name = os.path.realpath(argv[2])
            convert_all_ppt(directory_name)
    else:
        print("Wrong parameter. To get help, use: ppt2pdf.py -h")

if __name__ == '__main__':
    main(sys.argv)
