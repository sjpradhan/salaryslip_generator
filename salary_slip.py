import os
import re
import pandas as pd
import concurrent.futures
from reportlab.lib import colors
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4

def sanitize_filename(filename):
    """
    Sanitize the filename by removing special characters that aren't allowed in file names.
    """
    return re.sub(r'[\\/*?:"<>|]', "_", filename)

def format_number(value):
    """
    Helper function to format a number with thousand separators and 2 decimal places.
    If the value is NaN or None, returns an empty string.
    """
    if pd.notna(value):
        return "{:,.2f}".format(value)
    else:
        return ""

def generate_single_pdf(row, logo_path, output_dir):
    """
    Function to generate a single PDF for each employee.
    """
    # Sanitize the file name and construct the file path
    sanitized_name = sanitize_filename(f"{row['Month & Year']} {row['Employee Code']} {row['Name']}.pdf")
    file_path = os.path.join(output_dir, sanitized_name)

    # Create PDF canvas
    c = canvas.Canvas(file_path, pagesize=A4)
    width, height = A4

    # Adding some content to the PDF
    # Add company logo
    if logo_path:
        logo_width = 250
        logo_height = 190
        x_position = width - logo_width - 30
        y_position = height - 200
        c.drawImage(logo_path, x_position, y_position, width=logo_width, height=logo_height)

    month_year = str(row['Month & Year']) if pd.notna(row['Month & Year']) else ""
    employee_code = str(row['Employee Code']) if pd.notna(row['Employee Code']) else ""
    name = str(row['Name']) if pd.notna(row['Name']) else ""
    designation = str(row['Designation']) if pd.notna(row['Designation']) else ""
    date_of_joining = row['Date Of Joining'].strftime('%d-%m-%Y') if pd.notna(row['Date Of Joining']) else ""
    account_no = str(int(row['Account No'])) if pd.notna(row['Account No']) else ""
    bank_ifsc = str(row['Bank IFSC']) if pd.notna(row['Bank IFSC']) else ""
    pan_number = str(row['PAN No']) if pd.notna(row['PAN No']) else ""
    uan_number = str(int(row['UAN No'])) if pd.notna(row['UAN No']) else ""
    pf_number = str(row['PF No']) if pd.notna(row['PF No']) else ""
    esi_number = str(int(row['ESI No'])) if pd.notna(row['ESI No']) and row['ESI No'] != 0 else ""
    month_days = str(row['Month Days']) if pd.notna(row['Month Days']) else ""
    actual_payable_days = str(row['Actual Payable Days']) if pd.notna(row['Actual Payable Days']) else ""
    loss_of_pay_days = str(row['Loss Of Pay Days']) if pd.notna(row['Loss Of Pay Days']) else ""

    # Now apply the function to each column
    basic = format_number(row['Basic'])
    da = format_number(row['DA'])
    hra = format_number(row['HRA'])
    other_allowances = format_number(row['Other Allowance'])
    personal_pay = format_number(row['Personal Pay'])
    arrear = format_number(row['Arrear'])
    leave_encash = format_number(row['Leave Encash'])
    provident_fund = format_number(row['Provident Fund'])
    esi = format_number(row['ESI'])
    contribution_total = format_number(row['Contribution Total'])
    professional_tax = format_number(row['Professional Tax'])
    leave_without_pay = format_number(row['Leave without Pay'])
    tds = format_number(row['TDS'])
    gross_pay = format_number(row['Gross Pay'])
    gross_deduction = format_number(row['Gross Deduction'])
    net_salary_payable = format_number(row['Net Salary Payable (In Rs)'])

    # Define Font Style
    normal_font_style = "Helvetica"
    bold_font_style = "Helvetica-Bold"

    # Add title
    c.setFont(normal_font_style, 16)
    c.setStrokeColor(colors.black)  # Set border color
    c.drawString(50, height - 85, f"PAYSLIP " + str(month_year))

    # Add Address
    c.setFont(normal_font_style, 11)
    c.drawString(50, height - 100, "24X7 Moneyworks Consulting Private Limited")
    c.drawString(50, height - 135, "Registered Office: 205-206, Corner Point, Jetalpur Road")
    c.drawString(50, height - 150, "Development Location: 509, Midtown Complex, Jetalpur Road")
    c.drawString(50, height - 165, "Vadodara-390007")

    c.line(40, height - 190, width - 40, height - 190)  # 1st Horizontal line separator
    c.line(300, height - 190, 300, height - 300)  # 2nd Vertical line separator
    c.line(40, height - 300, width - 40, height - 300)  # 3rd Horizontal line separator
    c.line(40, height - 330, width - 40, height - 330)  # 4th Horizontal line separator

    c.setStrokeColorRGB(0.8, 0.8, 0.8)  # Light gray color (you can adjust the values for different shades)
    c.setLineWidth(0.7)  # Adjust line thickness (default is 1)
    c.line(40, height - 348, width - 40, height - 348)  # 5th Horizontal line separator

    c.setStrokeColorRGB(0, 0, 0)  # Default black color
    c.setLineWidth(1)  # Default line thickness
    c.line(40, height - 390, width - 40, height - 390)  # 6th Horizontal line separator
    c.line(170, height - 390, 170, height - 505)  # 7th Vertical line separator
    c.line(300, height - 390, 300, height - 505)  # 8th Vertical line separator
    c.line(425, height - 390, 425, height - 505)  # 9th Vertical line separator
    c.line(40, height - 505, width - 40, height - 505)  # 10th Horizontal line separator

    # Add Employee Details dynamically
    c.setFont(normal_font_style, 11)
    c.setFillColor(colors.black)

    # Draw text labels with aligned colons
    x_label = 50  # Starting position for the labels
    x_colon = 150  # Position for the colon symbol

    # Draw employee details
    c.drawString(x_label, height - 210, "Employee Code")
    c.drawString(x_colon, height - 210, ":  " + str(employee_code))
    c.drawString(x_label, height - 230, "Name")
    c.drawString(x_colon, height - 230, ":  " + name)
    c.drawString(x_label, height - 250, "Designation")
    c.drawString(x_colon, height - 250, ":  " + designation)
    c.drawString(x_label, height - 270, "Date Of Joining")
    c.drawString(x_colon, height - 270, ":  " + str(date_of_joining))
    c.drawString(x_label, height - 290, "Account No")
    c.drawString(x_colon, height - 290, ":  " + str(account_no))

    # Draw text labels with aligned colons for additional information (if needed)
    x_label = 310  # Starting position for the labels
    x_colon = 385  # Position for the colon symbol

    c.drawString(x_label, height - 210, "Bank IFSC")
    c.drawString(x_colon, height - 210, ":  " + str(bank_ifsc))
    c.drawString(x_label, height - 230, "PAN No")
    c.drawString(x_colon, height - 230, ":  " + str(pan_number))
    c.drawString(x_label, height - 250, "UAN No")
    c.drawString(x_colon, height - 250, ":  " + str(uan_number))
    c.drawString(x_label, height - 270, "PF No")
    c.drawString(x_colon, height - 270, ":  " + str(pf_number))
    c.drawString(x_label, height - 290, "ESI No")
    c.drawString(x_colon, height - 290, ":  " + str(esi_number))

    c.line(300, height - 190, 300, height - 300)  # 2nd Vertical line separator
    c.line(40, height - 300, width - 40, height - 300)  # 3rd Horizontal line separator

    # Add Salary Details
    c.setFont(bold_font_style, 14)
    c.setFillColor(colors.black)
    c.drawString(50, height - 318, "Salary Details")

    c.line(40, height - 330, width - 40, height - 330)  # 4th Horizontal line separator

    # Add Salary Details (static headers, can be updated dynamically as needed)
    c.setFont(bold_font_style, 9)
    c.setFillColor(colors.black)  # Set text color for contrast

    # Set the initial vertical position for the labels and data
    label_y_position = height - 342  # starting position for labels

    # Draw the labels
    c.drawString(75, label_y_position, "Month Days")
    c.drawString(250, label_y_position, "Actual Payable Days")
    c.drawString(420, label_y_position, "Loss Of Pay Days")

    # For the next line, move the vertical position down by a certain amount
    data_y_position = label_y_position - 17  # Decrease by 15 for next line
    c.setFont(normal_font_style, 10)

    # Draw the data values right below the corresponding labels
    c.drawString(90, data_y_position, str(month_days))  # Replace with the actual data value
    c.drawString(280, data_y_position, str(actual_payable_days))  # Replace with the actual data value
    c.drawString(450, data_y_position, str(loss_of_pay_days))  # Replace with the actual data value

    c.setStrokeColorRGB(0.8, 0.8, 0.8)  # Light gray color (you can adjust the values for different shades)
    c.setLineWidth(0.7)  # Adjust line thickness (default is 1)
    c.line(40, height - 348, width - 40, height - 348)  # 5th Horizontal line separator

    # Set font for Earning & Contribution section
    c.setFont(bold_font_style, 11)

    # Add Earning & Contribution
    c.drawString(75, height - 385, "EARNINGS")
    c.drawString(210, height - 385, "AMOUNT")
    c.drawString(320, height - 385, "CONTRIBUTION")
    c.drawString(465, height - 385, "AMOUNT")

    # Draw text labels with aligned colons for Earning & Contributions
    x_label = 50  # Starting position for the labels
    x_colon = 210  # Position for the colon symbol
    c.setFont(normal_font_style, 11)

    # Draw employee details
    c.drawString(x_label, height - 410, "Basic")
    c.drawString(x_colon, height - 410, " " + str(basic))
    c.drawString(x_label, height - 425, "DA")
    c.drawString(x_colon, height - 425, " " + str(da))
    c.drawString(x_label, height - 440, "HRA")
    c.drawString(x_colon, height - 440, " " + str(hra))
    c.drawString(x_label, height - 455, "Other Allowance")
    c.drawString(x_colon, height - 455, " " + str(other_allowances))

    # Initial starting height for the rows
    current_height = height - 470  # Position for "Personal Pay"

    # Draw "Personal Pay"
    c.drawString(x_label, current_height, "Personal Pay")
    c.drawString(x_colon, current_height, " " + str(personal_pay))

    # Check if Arrear is present and adjust height accordingly
    if arrear != "0" and arrear != "" and arrear is not None:
        current_height -= 15  # Move the height down for Arrear
        c.drawString(x_label, current_height, "Arrear")
        c.drawString(x_colon, current_height, " " + str(arrear))

    # Check if Leave Encash is present and adjust height accordingly
    if leave_encash != "0" and leave_encash != "" and leave_encash is not None:
        current_height -= 15  # Move the height down for Leave Encash
        c.drawString(x_label, current_height, "Leave Encash")
        c.drawString(x_colon, current_height, " " + str(leave_encash))

    c.drawString(x_label, height - 520, "GROSS PAY")
    c.drawString(x_colon, height - 520, " " + str(gross_pay))

    c.line(170, height - 390, 170, height - 505)  # 7th Vertical line separator
    c.line(300, height - 390, 300, height - 505)  # 8th Vertical line separator
    c.line(425, height - 390, 425, height - 505)  # 9th Vertical line separator
    c.line(40, height - 505, width - 40, height - 505)  # 10th Horizontal line separator

    x_label = 310  # Starting position for the labels
    x_colon = 460  # Position for the colon symbol

    c.drawString(x_label, height - 410, "Provident Fund")
    c.drawString(x_colon, height - 410, " " + str(provident_fund))
    if esi != "0" and esi != "" and esi is not None:
        c.drawString(x_label, height - 425, "ESI")
        c.drawString(x_colon, height - 425, " " + str(esi))
    c.setFont(bold_font_style, 11)
    c.drawString(x_label, height - 440, "Contribution Total")
    c.drawString(x_colon, height - 440, " " + str(contribution_total))
    c.drawString(x_label, height - 455, "Taxes & Deductions")
    c.setFont(normal_font_style, 11)
    c.drawString(x_label, height - 470, "Professional Tax")
    c.drawString(x_colon, height - 470, " " + str(professional_tax))
    c.drawString(x_label, height - 485, "Leave without Pay")
    c.drawString(x_colon, height - 485, " " + str(leave_without_pay))
    if tds != "0" and tds != "" and tds is not None:
        c.drawString(x_label, height - 500, "TDS")
        c.drawString(x_colon, height - 500, " " + str(tds))
    c.drawString(x_label, height - 520, "GROSS DEDUCTION")
    c.drawString(x_colon, height - 520, " " + str(gross_deduction))

    # Add box with colored background
    c.setFillColorRGB(0.9, 0.9, 0.9)  # Set background color for the box
    c.setStrokeColorRGB(0.9, 0.9, 0.9)  # No border around the rectangle
    c.rect(40, height - 600 - 100, width - 80, 80, fill=1)  # Draw a filled rectangle for the box

    # Add Net Salary
    c.setFont(bold_font_style, 11)
    c.setFillColor(colors.black)
    c.drawString(100, height - 665, "Net Salary Payable (In Rs)")
    c.drawString(390, height - 665, " " + net_salary_payable)
    c.setFont(normal_font_style, 11)
    c.drawString(100, height - 740, "**Note : This is a computer-generated document. No signature is required.")

    # Draw border around the content area (including the header)
    c.setStrokeColor(colors.black)  # Set border color
    c.setLineWidth(1)
    border_height = 730  # Customize the height of the border
    c.rect(40, height - 780, width - 80, border_height)  # Draw the border (with padding)

    # Finalize the PDF
    c.save()

def generate_salary_slip(excel_file, logo_path, output_dir):
    """
    Generate salary slips for all employees listed in the provided Excel file.
    """
    # Read data from the Excel file
    df = pd.read_excel(excel_file)

    # Create output directory if it does not exist
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)

    # Using ThreadPoolExecutor for parallel PDF generation
    with concurrent.futures.ThreadPoolExecutor(max_workers=4) as executor:
        futures = []
        for _, row in df.iterrows():
            futures.append(executor.submit(generate_single_pdf, row, logo_path, output_dir))

        # Wait for all threads to complete
        concurrent.futures.wait(futures)

    # Return paths to all generated PDFs
    return [
        os.path.join(output_dir, sanitize_filename(f"{row['Month & Year']} {row['Employee Code']} {row['Name']}.pdf"))
        for _, row in df.iterrows()]

