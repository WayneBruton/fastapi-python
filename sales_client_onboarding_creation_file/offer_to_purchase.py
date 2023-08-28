# import copy
# import os

from fpdf import FPDF, XPos, YPos
from sales_client_onboarding_creation_file.standard_conditions import create_standard_conditions
from sales_client_onboarding_creation_file.number_of_purchasers import create_purchaser_details
from sales_client_onboarding_creation_file.purchaser_signatory_details import create_signatories
from sales_client_onboarding_creation_file.annexure_list import create_annexure_list


class PDF(FPDF):

    def __init__(self, data):
        super().__init__()
        self.data = data
        self.unit = 'mm'
        self.format = 'A4'

        # make right margin bigger
        # self.r_margin = 25

    def footer(self):
        self.set_y(-15)
        self.set_font('helvetica', 'B', 6)

        self.cell(50, 10, f"{self.data['development'].upper()} Agreement of Sale", align="L")
        self.set_font('helvetica', 'I', 6)

        self.cell(90, 10, f"Page {self.page_no()}/{{nb}}", align="C")

        # add an "initial" block on the bottom right of the footer with a border around it
        # if self.page_no() > 1:
        self.set_font('helvetica', '', 8)
        # make the font color light gray
        self.set_text_color(211, 211, 211)
        self.cell(0, 10, f"INITIAL                 ", align="R")
        self.set_line_width(0.5)
        # make the rectangle lines light gray
        self.set_draw_color(211, 211, 211)

        self.rect(165, 282, 30, 8)

        # put a bold line above the footer
        self.set_draw_color(0, 0, 0)
        self.set_line_width(0.5)
        self.line(5, self.get_y() - 1, 200, self.get_y() - 1)


def print_otp_pdf(data):
    # pdf = PDF('P', 'mm', 'A4')
    pdf = PDF(data)


    pdf.set_auto_page_break(auto=True, margin=15)

    pdf.add_page()
    # pdf.set_font('helvetica', '', 10)
    pdf.set_fill_color(211, 211, 211)
    pdf.set_font('helvetica', '', 10)
    ############################################
    if 'section' not in data:
        data['section'] = 21

    door_no = data['opportunity_code'][-4:]

    if data['development'] == 'Heron View' or data['development'] == 'Heron Fields':
        company = 'HERON PROJECTS PROPRIETARY LIMITED'
        company_reg_no = 'REGISTRATION NUMBER 2020/495056/07'
        if data['development'] == 'Heron View':
            street_address = 'Crystal Way, Langeberg Ridge'
            estimated_levy = 'R 1 200.00'
        if data['development'] == 'Heron Fields':
            street_address = 'Cnr of Langeberg & Ridge Road, Langeberg Ridge'
            estimated_levy = 'R 1 200.00'
    elif data['development'] == 'Endulini':
        company = 'ENDULINI PROJECTS PROPRIETARY LIMITED'
        company_reg_no = 'REGISTRATION NUMBER 2020/495057/07'
        street_address = 'Cnr of Kruis Street & Crammix Street, Brackenfell'
        estimated_levy = 'R 1 200.00'
    ############################################

    pdf.cell(0, 5, f"**Section/Door No. {data['section']}/{door_no}**", align='R', fill=False, new_x=XPos.LMARGIN,
             new_y=YPos.NEXT, markdown=True, )
    pdf.cell(0, 5, f"**{data['development']}**", align='R', fill=False, new_x=XPos.LMARGIN,
             new_y=YPos.NEXT, markdown=True, )
    pdf.cell(0, 20, "", new_x=XPos.LMARGIN, new_y=YPos.NEXT)

    pdf.set_font('helvetica', '', 20)
    pdf.cell(0, 10, f"**DEED OF SALE**", new_x=XPos.LMARGIN, new_y=YPos.NEXT,
             markdown=True,
             align="C")
    pdf.set_font('helvetica', '', 10)
    pdf.cell(0, 5, f"**(Sectional Title Off Plan)**", new_x=XPos.LMARGIN, new_y=YPos.NEXT,
             markdown=True,
             align="C")
    pdf.cell(0, 10, "", new_x=XPos.LMARGIN, new_y=YPos.NEXT)
    pdf.cell(0, 10, "", new_x=XPos.LMARGIN, new_y=YPos.NEXT)
    pdf.set_font('helvetica', '', 30)
    pdf.cell(0, 10, f"**{data['development'].upper()}**", new_x=XPos.LMARGIN, new_y=YPos.NEXT, markdown=True,
             align="C")
    pdf.set_font('helvetica', '', 8)
    pdf.cell(0, 5, "", new_x=XPos.LMARGIN, new_y=YPos.NEXT)

    pdf.cell(0, 5, f"MADE AND ENTERED INTO BY AND BETWEEN:", new_x=XPos.LMARGIN, new_y=YPos.NEXT, markdown=True,
             align="C")

    pdf.set_font('helvetica', '', 15)
    pdf.cell(0, 5, "", new_x=XPos.LMARGIN, new_y=YPos.NEXT)
    pdf.cell(0, 5, "", new_x=XPos.LMARGIN, new_y=YPos.NEXT)

    pdf.cell(0, 10, f"**{company}**", new_x=XPos.LMARGIN, new_y=YPos.NEXT, markdown=True,
             align="C")
    pdf.cell(0, 10, f"**{company_reg_no}**", new_x=XPos.LMARGIN, new_y=YPos.NEXT, markdown=True,
             align="C")

    pdf.set_font('helvetica', '', 10)
    pdf.cell(0, 5, f"**(herein represented by CHARLES NIXON MORGAN**",
             new_x=XPos.LMARGIN, new_y=YPos.NEXT, markdown=True,
             align="C")
    pdf.cell(0, 5, f"**duly authorised hereto in terms of a Resolution)**",
             new_x=XPos.LMARGIN, new_y=YPos.NEXT, markdown=True,
             align="C")
    pdf.cell(0, 20, f'("the Seller")',
             new_x=XPos.LMARGIN, new_y=YPos.NEXT, markdown=True,
             align="C")
    pdf.cell(0, 20, f'And',
             new_x=XPos.LMARGIN, new_y=YPos.NEXT, markdown=True,
             align="C")

    ############################################
    purchasers = []
    number_of_purchasers = int(data['opportunity_client_no'].split(" ")[0])

    if data['opportunity_client_type'] == 'Company':
        purchasers.append(data['opportunity_companyname'])
    else:
        if number_of_purchasers >= 1:
            purchasers.append(data['opportunity_firstname'] + " " + data['opportunity_lastname'])
        if number_of_purchasers >= 2:
            purchasers.append(data['opportunity_firstname_sec'] + " " + data['opportunity_lastname_sec'])
        if number_of_purchasers >= 3:
            purchasers.append(data['opportunity_firstname_3rd'] + " " + data['opportunity_lastname_3rd'])
        if number_of_purchasers >= 4:
            purchasers.append(data['opportunity_firstname_4th'] + " " + data['opportunity_lastname_4th'])
        if number_of_purchasers >= 5:
            purchasers.append(data['opportunity_firstname_5th'] + " " + data['opportunity_lastname_5th'])
        if number_of_purchasers >= 6:
            purchasers.append(data['opportunity_firstname_6th'] + " " + data['opportunity_lastname_6th'])
        if number_of_purchasers >= 7:
            purchasers.append(data['opportunity_firstname_7th'] + " " + data['opportunity_lastname_7th'])
        if number_of_purchasers >= 8:
            purchasers.append(data['opportunity_firstname_8th'] + " " + data['opportunity_lastname_8th'])
        if number_of_purchasers >= 9:
            purchasers.append(data['opportunity_firstname_9th'] + " " + data['opportunity_lastname_9th'])
        if number_of_purchasers >= 10:
            purchasers.append(data['opportunity_firstname_10th'] + " " + data['opportunity_lastname_10th'])

    ############################################

    for purchaser in purchasers:
        pdf.cell(0, 10, f"{purchaser.upper()}", new_x=XPos.LMARGIN, new_y=YPos.NEXT, markdown=True,
                 align="C")

    pdf.cell(0, 5, "", new_x=XPos.LMARGIN, new_y=YPos.NEXT)

    pdf.cell(0, 10,
             f"whose full particulars appear in the Information Schedule, forming an integral part of this Agreement.",
             new_x=XPos.LMARGIN, new_y=YPos.NEXT, markdown=True, align="C")
    pdf.cell(0, 20, f'("the Purchaser")',
             new_x=XPos.LMARGIN, new_y=YPos.NEXT, markdown=True,
             align="C")

    pdf.add_page()

    pdf.multi_cell(0, 5,
                   f'The Seller and the Purchaser mentioned in the Information Schedule hereby enter into an Agreement '
                   f'of Sale for the sale of the Property with a Sectional Title Unit thereon as described in C1 of '
                   f'the Information Schedule for the purchase price recorded in Clause E of the Information Schedule '
                   f'and on the terms set forth in the Information Schedule and Standard Terms and Conditions forming '
                   f'pages 7 to 30 hereof and Annexures **A - G**.',

                   new_x=XPos.LMARGIN, new_y=YPos.NEXT, markdown=True,
                   align="J")
    pdf.cell(0, 20, f'SIGNED AT_________________________THIS__________DAY OF___________________20_____',
             markdown=True,
             align="L", new_x=XPos.LMARGIN, new_y=YPos.NEXT)
    pdf.cell(0, 20, f'AS WITNESSES:',
             markdown=True,
             align="L", new_x=XPos.LMARGIN, new_y=YPos.NEXT)
    pdf.cell(90, 10, f'1.  _________________________',
             markdown=True,
             align="L", new_x=XPos.LMARGIN, new_y=YPos.NEXT)
    pdf.cell(90, 10, f'Name:  ______________________',
             markdown=True,
             align="L")
    pdf.cell(60, 10, f'___________________________________',
             markdown=True,
             align="L", new_x=XPos.LMARGIN, new_y=YPos.NEXT)
    pdf.cell(100, 10, f'ID:  _________________________',
             markdown=True,
             align="L")
    pdf.multi_cell(60, 5, f'for and behalf of the SELLER, the signatory warrants his/her authority hereto.',
                   markdown=True,
                   align="J", new_x=XPos.LMARGIN, new_y=YPos.NEXT)

    pdf.cell(0, 5, "", new_x=XPos.LMARGIN, new_y=YPos.NEXT)

    purchaser_signing_details = create_signatories(data, number_of_purchasers)

    signatories = 1

    for purchaser in purchaser_signing_details:
        signatories += 1
        pdf.cell(0, 20, f'SIGNED AT_________________________THIS__________DAY OF___________________20_____',
                 markdown=True,
                 align="L", new_x=XPos.LMARGIN, new_y=YPos.NEXT)

        pdf.cell(0, 20, f'AS WITNESSES:',
                 markdown=True,
                 align="L", new_x=XPos.LMARGIN, new_y=YPos.NEXT)
        pdf.cell(90, 10, f'1.  _________________________',
                 markdown=True,
                 align="L", new_x=XPos.LMARGIN, new_y=YPos.NEXT)
        pdf.cell(90, 10, f'Name:  ______________________',
                 markdown=True,
                 align="L")
        pdf.cell(60, 10, f'___________________________________',
                 markdown=True,
                 align="L", new_x=XPos.LMARGIN, new_y=YPos.NEXT)
        pdf.cell(100, 10, f'ID:  _________________________',
                 markdown=True,
                 align="L")
        pdf.multi_cell(60, 5,
                       f'{purchaser["purchaser"]}, the signatory warrants his/her authority hereto, where applicable',
                       markdown=True,
                       align="J", new_x=XPos.LMARGIN, new_y=YPos.NEXT)

        if signatories % 3 == 0:
            pdf.add_page()

        if purchaser["marital_status"] == "married-cop":
            signatories += 1
            pdf.cell(0, 20, f'SIGNED AT_________________________THIS__________DAY OF___________________20_____',
                     markdown=True,
                     align="L", new_x=XPos.LMARGIN, new_y=YPos.NEXT)

            pdf.cell(0, 20, f'AS WITNESSES:',
                     markdown=True,
                     align="L", new_x=XPos.LMARGIN, new_y=YPos.NEXT)
            pdf.cell(90, 10, f'1.  _________________________',
                     markdown=True,
                     align="L", new_x=XPos.LMARGIN, new_y=YPos.NEXT)
            pdf.cell(90, 10, f'Name:  ______________________',
                     markdown=True,
                     align="L")
            pdf.cell(60, 10, f'___________________________________',
                     markdown=True,
                     align="L", new_x=XPos.LMARGIN, new_y=YPos.NEXT)
            pdf.cell(100, 10, f'ID:  _________________________',
                     markdown=True,
                     align="L")
            pdf.multi_cell(60, 5,
                           f'{purchaser["info"]}',
                           markdown=True,
                           align="J", new_x=XPos.LMARGIN, new_y=YPos.NEXT)

            if signatories % 3 == 0:
                pdf.add_page()

        pdf.cell(0, 5, "", new_x=XPos.LMARGIN, new_y=YPos.NEXT)

    pdf.cell(0, 20, f'SIGNED AT_________________________THIS__________DAY OF___________________20_____',
             markdown=True,
             align="L", new_x=XPos.LMARGIN, new_y=YPos.NEXT)

    pdf.cell(0, 20, f'AS WITNESSES:',
             markdown=True,
             align="L", new_x=XPos.LMARGIN, new_y=YPos.NEXT)
    pdf.cell(90, 10, f'1.  _________________________',
             markdown=True,
             align="L", new_x=XPos.LMARGIN, new_y=YPos.NEXT)
    pdf.cell(90, 10, f'Name:  ______________________',
             markdown=True,
             align="L")
    pdf.cell(60, 10, f'___________________________________',
             markdown=True,
             align="L", new_x=XPos.LMARGIN, new_y=YPos.NEXT)
    pdf.cell(100, 10, f'ID:  _________________________',
             markdown=True,
             align="L")
    pdf.multi_cell(60, 5, f'for and on behalf of the **CONTRACTOR**, the signatory warrants his authority hereto in '
                          f'respect of the provisions of clause 18 of the Agreement.',
                   markdown=True,
                   align="J", new_x=XPos.LMARGIN, new_y=YPos.NEXT)

    pdf.cell(0, 5, "", new_x=XPos.LMARGIN, new_y=YPos.NEXT)

    pdf.add_page()


    pdf.cell(0, 10, "**INFORMATION SCHEDULE**", align="C", markdown=True, new_x=XPos.LMARGIN, new_y=YPos.NEXT)

    pdf.cell(20, 7, "**A**", align="C", markdown=True, border=1)
    pdf.cell(50, 7, "**Seller**", align="L", markdown=True, border=1)
    pdf.cell(120, 7, "**Details**", align="C", markdown=True, new_x=XPos.LMARGIN, new_y=YPos.NEXT, border=1)

    pdf.cell(20, 7, "A1", align="C", markdown=True, border=1)
    pdf.cell(50, 7, "Full Name", align="L", markdown=True, border=1)
    pdf.cell(120, 7, f"**{company}**", align="L", markdown=True, new_x=XPos.LMARGIN, new_y=YPos.NEXT, border=1)

    pdf.cell(20, 7, "A2", align="C", markdown=True, border=1)
    pdf.cell(50, 7, "Registration Number", align="L", markdown=True, border=1)
    pdf.cell(120, 7, f"**{company_reg_no.split('REGISTRATION NUMBER ')[1]}**", align="L", markdown=True,
             new_x=XPos.LMARGIN, new_y=YPos.NEXT, border=1)

    pdf.cell(20, 10, "A3", align="C", markdown=True, border=1)
    pdf.cell(50, 10, "Address (Street & Postal)", align="L", markdown=True, border=1)
    pdf.multi_cell(120, 5,
                   f"**Office 2, First Floor 251 Durban Rd, Bo-Oakdale, Bellville, 7530 & PO Box 1807 Bellville 7536**",
                   align="J", markdown=True,
                   new_x=XPos.LMARGIN, new_y=YPos.NEXT, border=1)

    pdf.cell(20, 7, "A4", align="C", markdown=True, border=1)
    pdf.cell(50, 7, "Telephone", align="L", markdown=True, border=1)
    pdf.cell(120, 7, f"**021 919 9944**", align="L", markdown=True,
             new_x=XPos.LMARGIN, new_y=YPos.NEXT, border=1)

    pdf.cell(20, 7, "A5", align="C", markdown=True, border=1)
    pdf.cell(50, 7, "E-mail", align="L", markdown=True, border=1)
    # set font color to blue
    pdf.set_text_color(0, 0, 255)
    pdf.cell(120, 7, f"**izolda@opportunity.co.za**", align="L", markdown=True,
             new_x=XPos.LMARGIN, new_y=YPos.NEXT, border=1)
    # set font color back to black
    pdf.set_text_color(0, 0, 0)

    pdf.cell(0, 5, "", new_x=XPos.LMARGIN, new_y=YPos.NEXT)

    ##########################################
    # Purchaser Details
    if data['opportunity_client_type'] == 'Company':
        if data['opportunity_legal_type'] == 'Company':
            b1_label = 'Company Name'
            b1_value = data['opportunity_companyname']
            b2_label = 'Company Registration Number'
            b2_value = data['opportunity_companyregistrationNo']
            b6_value = "Director"

        else:
            b1_label = 'Trust Name'
            b1_value = data['opportunity_companyname']
            b2_label = 'Trust Number'
            b2_value = data['opportunity_companyregistrationNo']
            b6_value = "Trustee"

        pdf.cell(20, 7, "**B**", align="C", markdown=True, border=1)
        pdf.cell(50, 7, "**Purchaser**", align="L", markdown=True, border=1)
        pdf.cell(120, 7, "**Details**", align="C", markdown=True, new_x=XPos.LMARGIN, new_y=YPos.NEXT, border=1)

        pdf.cell(20, 7, "B1", align="C", markdown=True, border=1)
        pdf.cell(50, 7, f"{b1_label}", align="L", markdown=True, border=1)
        pdf.cell(120, 7, f"**{b1_value}**", align="L", markdown=True,
                 new_x=XPos.LMARGIN, new_y=YPos.NEXT, border=1)

        pdf.cell(20, 7, "B2", align="C", markdown=True, border=1)
        pdf.cell(50, 7, f"{b2_label}", align="L", markdown=True, border=1)
        pdf.cell(120, 7, f"**{b2_value}**", align="L", markdown=True,
                 new_x=XPos.LMARGIN, new_y=YPos.NEXT, border=1)

        pdf.cell(20, 10, "B3", align="C", markdown=True, border=1)
        pdf.cell(50, 10, "Address (Street & Postal)", align="L", markdown=True, border=1)
        pdf.multi_cell(120, 5, f"**{data['opportunity_company_residential_address']}** **&** "
                               f"**{data['opportunity_company_postal_address']}**", align="J", markdown=True,
                       new_x=XPos.LMARGIN, new_y=YPos.NEXT, border=1)

        pdf.cell(20, 7, "B4", align="C", markdown=True, border=1)
        pdf.cell(50, 7, "Telephone", align="L", markdown=True, border=1)
        pdf.cell(120, 7, f"**{data['opportunity_company_contact']}**", align="L", markdown=True,
                 new_x=XPos.LMARGIN, new_y=YPos.NEXT, border=1)

        pdf.cell(20, 7, "B5", align="C", markdown=True, border=1)
        pdf.cell(50, 7, "E-mail", align="L", markdown=True, border=1)
        pdf.set_text_color(0, 0, 255)
        pdf.cell(120, 7, f"**{data['opportunity_company_email']}**", align="L", markdown=True,
                 new_x=XPos.LMARGIN, new_y=YPos.NEXT, border=1)
        pdf.set_text_color(0, 0, 0)

        pdf.cell(20, 7, "B6", align="C", border=1)
        pdf.cell(50, 7, "Signatory for Purchaser", align="L", markdown=True, border=1)
        pdf.cell(120, 7,
                 f"**{data['opportunity_firstname']} {data['opportunity_lastname']} - (Resolution/s to be attached)**",
                 align="L", markdown=True,
                 new_x=XPos.LMARGIN, new_y=YPos.NEXT, border=1)

        pdf.cell(20, 7, "B7", align="C", markdown=True, border=1)
        pdf.cell(50, 7, "Capacity", align="L", markdown=True, border=1)
        pdf.cell(120, 7, f"**{b6_value}**", align="L",
                 markdown=True,
                 new_x=XPos.LMARGIN, new_y=YPos.NEXT, border=1)

    else:
        purchaser_details = create_purchaser_details(data, number_of_purchasers)

        # Add the purchaser details to the pdf
        for purchaser in purchaser_details:
            if purchaser['value_b4'] == 'notMarried':
                purchaser['value_b4'] = 'Not Married'
            pdf.cell(20, 7, f"**B**", align="C", markdown=True, border=1)
            pdf.cell(50, 7, f"**{purchaser['header']}**", align="L", markdown=True, border=1)
            pdf.cell(120, 7, "**Details**", align="C", markdown=True, new_x=XPos.LMARGIN, new_y=YPos.NEXT, border=1)

            pdf.cell(20, 7, "B1", align="C", markdown=True, border=1)
            pdf.cell(50, 7, f"{purchaser['label_b1']}", align="L", markdown=True, border=1)
            pdf.cell(120, 7, f"**{purchaser['value_b1']}**", align="L", markdown=True,
                     new_x=XPos.LMARGIN, new_y=YPos.NEXT, border=1)

            pdf.cell(20, 7, "B2", align="C", markdown=True, border=1)
            pdf.cell(50, 7, f"{purchaser['label_b2']}", align="L", markdown=True, border=1)
            pdf.cell(120, 7, f"**{purchaser['value_b2']}**", align="L", markdown=True,
                     new_x=XPos.LMARGIN, new_y=YPos.NEXT, border=1)

            pdf.cell(20, 10, "B3", align="C", markdown=True, border=1)
            pdf.cell(50, 10, f"{purchaser['label_b3']}", align="L", markdown=True, border=1)
            pdf.multi_cell(120, 5, f"**{purchaser['value_b3']}**", align="J", markdown=True,
                           new_x=XPos.LMARGIN, new_y=YPos.NEXT, border=1)

            pdf.cell(20, 7, "B4", align="C", markdown=True, border=1)
            pdf.cell(50, 7, f"{purchaser['label_b4']}", align="L", markdown=True, border=1)
            pdf.cell(120, 7, f"**{purchaser['value_b4']}**", align="L", markdown=True,
                     new_x=XPos.LMARGIN, new_y=YPos.NEXT, border=1)

            pdf.cell(20, 7, "B5", align="C", markdown=True, border=1)
            pdf.cell(50, 7, f"{purchaser['label_b5']}", align="L", markdown=True, border=1)
            pdf.cell(120, 7, f"**{purchaser['value_b5']}**", align="L", markdown=True,
                     new_x=XPos.LMARGIN, new_y=YPos.NEXT, border=1)

            pdf.cell(20, 7, "B6", align="C", markdown=True, border=1)
            pdf.cell(50, 7, f"{purchaser['label_b6']}", align="L", markdown=True, border=1)
            pdf.cell(120, 7, f"**{purchaser['value_b6']}**", align="L", markdown=True,
                     new_x=XPos.LMARGIN, new_y=YPos.NEXT, border=1)

            pdf.cell(20, 7, "B7", align="C", markdown=True, border=1)
            pdf.cell(50, 7, f"{purchaser['label_b7']}", align="L", markdown=True, border=1)
            pdf.set_text_color(0, 0, 255)
            pdf.cell(120, 7, f"**{purchaser['value_b7']}**", align="L", markdown=True,
                     new_x=XPos.LMARGIN, new_y=YPos.NEXT, border=1)
            pdf.set_text_color(0, 0, 0)

            pdf.cell(0, 5, "", new_x=XPos.LMARGIN, new_y=YPos.NEXT)

    # pdf.add_page()

    pdf.cell(20, 7, f"**C**", align="C", markdown=True, border=1)
    pdf.cell(50, 7, f"**The Property**", align="L", markdown=True, border=1)
    pdf.cell(120, 7, "", align="C", markdown=True, new_x=XPos.LMARGIN, new_y=YPos.NEXT, border=1)

    pdf.cell(20, 15, f"C1", align="C", markdown=True, border=1)
    pdf.cell(50, 15, f"Unit", align="L", markdown=True, border=1)
    pdf.multi_cell(120, 5,
                   f"**Unit No:** {door_no} having an approximate floor area of .......... Square Metre's as "
                   f"reflected on the Development and Unit Plans annexed hereto (marked \"**A & B**\")",
                   align="J", markdown=True, new_x=XPos.LMARGIN, new_y=YPos.NEXT, border=1)

    ##############################################

    garden_number = 0
    garden_size = 0

    if data['opportunity_gardenNumber'] != 0 and data['opportunity_gardenNumber'] is not None:
        garden_number = data['opportunity_gardenNumber']
        garden_size = data['opportunity_gardenSize']

    aditional_parking_bays = data['opportunity_additional_bay']
    allocated_parking_bay = data['opportunity_originalBayNo']
    # second_parking_bay = data['opportunity_parking_base']
    second_parking_bay = data.get('opportunity_parking_base', "")
    # third_parking_bay = data['opportunity_parking_base2']
    third_parking_bay = data.get('opportunity_parking_base2', "")
    line_item = 1

    if garden_number != "0":
        pdf.cell(20, 20, f"C2", align="C", markdown=True, border="LTR")
        pdf.cell(50, 20, f"Exclusive Use Area's", align="L", markdown=True, border="LTR")
        pdf.multi_cell(120, 5,
                       f"**{line_item}. Garden {garden_number}** having an area of approximately {garden_size} "
                       f"square metres and shall be allocated to the Purchaser in terms of Section 27A of the "
                       f"Sectional Titles Act.  The Garden is indicated on the Exclusive Use Area Plan annexed hereto "
                       f"(marked **\"A\")**.",
                       align="J", markdown=True, new_x=XPos.LMARGIN, new_y=YPos.NEXT, border="LTR")

        line_item += 1
        if aditional_parking_bays != 0:
            border = "LR"
        else:
            border = "LBR"

        pdf.cell(20, 20, f"", align="C", markdown=True, border=border)
        pdf.cell(50, 20, f"", align="L", markdown=True, border=border)
        pdf.multi_cell(120, 5,
                       f"**{line_item}. Parking {allocated_parking_bay}** having an area of approximately .........."
                       f".. square metres and shall be allocated to the Purchaser in terms of Section 27A of the "
                       f"Sectional Titles Act.  The Parking is indicated on the Exclusive Use Area Plan annexed "
                       f"hereto (marked **\"A\"**).",
                       align="J", markdown=True, new_x=XPos.LMARGIN, new_y=YPos.NEXT, border=border)

        line_item += 1
        if aditional_parking_bays == 1:
            border = "LBR"

        elif aditional_parking_bays == 2:
            border = "LR"

        if aditional_parking_bays >= 1:
            pdf.cell(20, 20, f"", align="C", markdown=True, border=border)
            pdf.cell(50, 20, f"", align="L", markdown=True, border=border)
            pdf.multi_cell(120, 5,
                           f"**{line_item}. Additional Parking {second_parking_bay}** having an area of approximately"
                           f" ............ square metres and shall be allocated to the Purchaser in terms of Section "
                           f"27A of the Sectional Titles Act.  The Parking is indicated on the Exclusive Use Area "
                           f"Plan annexed hereto (marked **\"A\"**).",
                           align="J", markdown=True, new_x=XPos.LMARGIN, new_y=YPos.NEXT, border=border)

        if aditional_parking_bays == 2:
            line_item += 1
            border = "LBR"
            pdf.cell(20, 20, f"", align="C", markdown=True, border=border)
            pdf.cell(50, 20, f"", align="L", markdown=True, border=border)
            pdf.multi_cell(120, 5,
                           f"**{line_item}. Additional Parking {third_parking_bay}** having an area of approximately "
                           f"............ square metres and shall be allocated to the Purchaser in terms of Section "
                           f"27A of the Sectional Titles Act.  The Parking is indicated on the Exclusive Use Area "
                           f"Plan annexed hereto (marked **\"A\"**).",
                           align="J", markdown=True, new_x=XPos.LMARGIN, new_y=YPos.NEXT, border=border)

    else:
        line_item = 1
        if aditional_parking_bays != 0:
            border = "LR"
        else:
            border = "LBR"

        pdf.cell(20, 20, f"C2", align="C", markdown=True, border=border)
        pdf.cell(50, 20, f"Exclusive Use Area's", align="L", markdown=True, border=border)
        pdf.multi_cell(120, 5,
                       f"**{line_item}. Parking {allocated_parking_bay}** having an area of approximately .........."
                       f".. square metres and shall be allocated to the Purchaser in terms of Section 27A of the "
                       f"Sectional Titles Act.  The Parking is indicated on the Exclusive Use Area Plan annexed "
                       f"hereto (marked **\"A\"**).",
                       align="J", markdown=True, new_x=XPos.LMARGIN, new_y=YPos.NEXT, border=border)

        line_item += 1
        if aditional_parking_bays == 1:
            border = "LBR"

        elif aditional_parking_bays == 2:
            border = "LR"

        if aditional_parking_bays >= 1:
            pdf.cell(20, 20, f"", align="C", markdown=True, border=border)
            pdf.cell(50, 20, f"", align="L", markdown=True, border=border)
            pdf.multi_cell(120, 5,
                           f"**{line_item}. Additional Parking {second_parking_bay}** having an area of approximately "
                           f"............ square metres and shall be allocated to the Purchaser in terms of Section "
                           f"27A of the Sectional Titles Act.  The Parking is indicated on the Exclusive Use Area "
                           f"Plan annexed hereto (marked **\"A\"**).",
                           align="J", markdown=True, new_x=XPos.LMARGIN, new_y=YPos.NEXT, border=border)

        if aditional_parking_bays == 2:
            line_item += 1
            border = "LBR"
            pdf.cell(20, 20, f"", align="C", markdown=True, border=border)
            pdf.cell(50, 20, f"", align="L", markdown=True, border=border)
            pdf.multi_cell(120, 5,
                           f"**{line_item}. Additional Parking {third_parking_bay}** having an area of approximately "
                           f"............ square metres and shall be allocated to the Purchaser in terms of Section "
                           f"27A of the Sectional Titles Act.  The Parking is indicated on the Exclusive Use Area "
                           f"Plan annexed hereto (marked **\"A\"**).",
                           align="J", markdown=True, new_x=XPos.LMARGIN, new_y=YPos.NEXT, border=border)

    ##############################################

    pdf.cell(20, 10, f"C3", align="C", markdown=True, border=1)
    pdf.cell(50, 10, f"Street Address", align="L", markdown=True, border=1)
    pdf.multi_cell(120, 5,
                   f"**Unit No. {door_no}, {data['development']}, {street_address}**",
                   align="J", markdown=True, new_x=XPos.LMARGIN, new_y=YPos.NEXT, border=1)

    pdf.cell(0, 3, f"", align="C", markdown=True, border=0, new_x=XPos.LMARGIN, new_y=YPos.NEXT)

    pdf.cell(20, 8, f"**D**", align="C", markdown=True, border=1)
    pdf.cell(50, 8, f"**Estimated Occupation**", align="L", markdown=True, border=1)
    pdf.multi_cell(120, 4,
                   f"**As per and subject to Clause 11.13 of the Agreement Estimated to be on or around:............",
                   align="J", markdown=True, new_x=XPos.LMARGIN, new_y=YPos.NEXT, border=1)

    pdf.cell(0, 3, f"", align="C", markdown=True, border=0, new_x=XPos.LMARGIN, new_y=YPos.NEXT)

    pdf.cell(20, 7, f"**E**", align="C", markdown=True, border=1)
    pdf.cell(50, 7, f"**Purchase Price**", align="L", markdown=True, border=1)
    pdf.cell(120, 7, "", align="C", markdown=True, new_x=XPos.LMARGIN, new_y=YPos.NEXT, border=1)

    pdf.cell(20, 7, "E1", align="C", markdown=True, border=1)
    pdf.cell(50, 7, f"Purchase Price (VAT Inclusive)", align="L", markdown=True, border=1)
    pdf.cell(120, 7, f"   **R {data['opportunity_contract_price']:,.2f}**", align="L", markdown=True,
             new_x=XPos.LMARGIN, new_y=YPos.NEXT, border=1)

    pdf.cell(20, 10, "E2", align="C", markdown=True, border=1)
    pdf.cell(50, 10, f"Deposit", align="L", markdown=True, border=1)
    pdf.multi_cell(120, 5, f"**R {float(data['opportunity_deposite']):,.2f} (Note: Minimum of R30,000.00 deposit "
                           f"payable within 3 days of Signature Date to secure)**", align="J", markdown=True,
                   new_x=XPos.LMARGIN, new_y=YPos.NEXT, border=1)

    pdf.cell(20, 7, "E3", align="C", markdown=True, border=1)
    pdf.cell(50, 7, f"Bond or Balance if Cash", align="L", markdown=True, border=1)
    pdf.cell(120, 7, f"   **R**", align="L", markdown=True,
             new_x=XPos.LMARGIN, new_y=YPos.NEXT, border=1)

    pdf.cell(0, 3, f"", align="C", markdown=True, border=0, new_x=XPos.LMARGIN, new_y=YPos.NEXT)

    pdf.cell(20, 15, "**F**", align="C", markdown=True, border=1)
    pdf.cell(50, 15, f"**Bond Costs**", align="L", markdown=True, border=1)
    pdf.multi_cell(120, 5, f"**The Purchaser will be liable for payment of initiation and/or valuation "
                           f"(bank administration) fees as may be charged by the bank, and as further set out herein**",
                   align="J", markdown=True, new_x=XPos.LMARGIN, new_y=YPos.NEXT, border=1)

    pdf.cell(0, 3, f"", align="C", markdown=True, border=0, new_x=XPos.LMARGIN, new_y=YPos.NEXT)

    pdf.cell(20, 7, f"**G**", align="C", markdown=True, border=1)
    pdf.cell(50, 7, f"**Selling Details**", align="L", markdown=True, border=1)
    pdf.cell(120, 7, "", align="C", markdown=True, new_x=XPos.LMARGIN, new_y=YPos.NEXT, border=1)

    pdf.cell(20, 7, "G1", align="C", markdown=True, border=1)
    pdf.cell(50, 7, f"Selling Agency", align="L", markdown=True, border=1)
    pdf.cell(120, 7, f"**OPPORTUNITY GLOBAL INVESTMENT PROPERTY (PTY) LTD**", align="L", markdown=True,
             new_x=XPos.LMARGIN, new_y=YPos.NEXT, border=1)

    pdf.cell(20, 7, "G2", align="C", markdown=True, border=1)
    pdf.cell(50, 7, f"Sales Agent", align="L", markdown=True, border=1)
    pdf.cell(120, 7, f"**{data['opportunity_sale_agreement']}**", align="L", markdown=True,
             new_x=XPos.LMARGIN, new_y=YPos.NEXT, border=1)

    pdf.cell(0, 3, f"", align="C", markdown=True, border=0, new_x=XPos.LMARGIN, new_y=YPos.NEXT)

    pdf.cell(20, 7, f"**H**", align="C", markdown=True, border=1)
    pdf.cell(50, 7, f"**Estimated Monthly Levy**", align="L", markdown=True, border=1)
    pdf.cell(120, 7, f"**TBC By Managing Agent (Payable by Purchaser)** Approx {estimated_levy}", align="L",
             markdown=True, new_x=XPos.LMARGIN, new_y=YPos.NEXT, border=1)

    pdf.cell(0, 3, f"", align="C", markdown=True, border=0, new_x=XPos.LMARGIN, new_y=YPos.NEXT)

    if data['development'] == "Heron View":
        pdf.cell(20, 7, f"**I**", align="C", markdown=True, border=1)
        pdf.cell(50, 7, f"**Electricity Deposit**", align="L", markdown=True, border=1)
        pdf.cell(120, 7, f"**R1200.00 Payable upon Registration to the Transfer Attorneys", align="L",
                 markdown=True, new_x=XPos.LMARGIN, new_y=YPos.NEXT, border=1)

    pdf.cell(0, 3, f"", align="C", markdown=True, border=0, new_x=XPos.LMARGIN, new_y=YPos.NEXT)

    if data['development'] != "Endulini":
        pdf.cell(20, 7, f"**J**", align="C", markdown=True, border=1)
    else:
        pdf.cell(20, 7, f"**I**", align="C", markdown=True, border=1)
    pdf.cell(50, 7, f"**Mortgage Originator**", align="L", markdown=True, border=1)
    pdf.cell(120, 7, f"{data['opportunity_bond_originator']}", align="L",
             markdown=True, new_x=XPos.LMARGIN, new_y=YPos.NEXT, border=1)

    # extras

    pdf.cell(0, 5, f"", align="C", markdown=True, border=0, new_x=XPos.LMARGIN, new_y=YPos.NEXT)

    # make font larger
    pdf.set_font('helvetica', '', 14)

    pdf.cell(0, 7, f"**Additional Conditions & Notes:**", align="L", markdown=True, border=0, new_x=XPos.LMARGIN,
             new_y=YPos.NEXT)

    pdf.cell(0, 5, f"", align="C", markdown=True, border=0, new_x=XPos.LMARGIN, new_y=YPos.NEXT)

    # Revert font size
    pdf.set_font('helvetica', '', 11)

    pdf.cell(0, 7, f"**Inclusions and choices:**", align="L", markdown=True, border=0, new_x=XPos.LMARGIN,
             new_y=YPos.NEXT)

    ###################################

    if len(data['opportunity_specials']) > 0:
        for special in data['opportunity_specials']:
            pdf.cell(0, 7, f"{special}", align="L", markdown=True, border=0, new_x=XPos.LMARGIN,
                     new_y=YPos.NEXT)

    if data['opportunity_stove_option'] == 'Gas':
        pdf.cell(0, 7, f"Gas Stove Option chosen", align="L", markdown=True, border=0, new_x=XPos.LMARGIN,
                 new_y=YPos.NEXT)

    pdf.set_font('helvetica', '', 8)
    # make the font color light gray
    pdf.set_text_color(211, 211, 211)

    pdf.cell(0, 10, f"INITIAL                 ", align="R")
    # get the current x and y position

    pdf.set_line_width(0.5)
    # make the rectangle lines light gray
    pdf.set_draw_color(211, 211, 211)
    # draw the rectangle for the initials
    x = pdf.get_x() - 35
    y = pdf.get_y() + 1

    # draw a rectangle
    pdf.rect(x, y, 30, 10)

    # pdf.rect(165, 282, 30, 10)
    pdf.set_text_color(0, 0, 0)
    pdf.set_font('helvetica', '', 9)

    # print(pdf.get_y())
    if pdf.get_y() > 200:
        pdf.add_page()
    pdf.cell(0, 7, f"**Notes:**", align="L", markdown=True, border=0, new_x=XPos.LMARGIN,
             new_y=YPos.NEXT)

    if data['opportunity_notes'] != "":
        pdf.cell(0, 7, f"{data['opportunity_notes']}", align="L", markdown=True, border="B", new_x=XPos.LMARGIN,
                 new_y=YPos.NEXT)

    for x in range(0, 5):
        pdf.cell(0, 7, f"", align="L", markdown=True, border="B", new_x=XPos.LMARGIN,
                 new_y=YPos.NEXT)

    pdf.set_font('helvetica', '', 8)
    # make the font color light gray
    pdf.set_text_color(211, 211, 211)

    pdf.cell(0, 10, f"INITIAL                 ", align="R")
    # get the current x and y position

    pdf.set_line_width(0.5)
    # make the rectangle lines light gray
    pdf.set_draw_color(211, 211, 211)
    # draw the rectangle for the initials
    x = pdf.get_x() - 35
    y = pdf.get_y() + 1

    # draw a rectangle
    pdf.rect(x, y, 30, 10)

    # pdf.rect(165, 282, 30, 10)
    pdf.set_text_color(0, 0, 0)
    pdf.set_font('helvetica', '', 9)

    pdf.cell(0, 7, f"", align="C", markdown=True, border=0, new_x=XPos.LMARGIN, new_y=YPos.NEXT)
    pdf.cell(0, 7, f"", align="C", markdown=True, border=0, new_x=XPos.LMARGIN, new_y=YPos.NEXT)
    pdf.cell(0, 7, f"**Summary of Annexures**", align="C", markdown=True, border=0, new_x=XPos.LMARGIN, new_y=YPos.NEXT)
    pdf.cell(0, 7, f"", align="C", markdown=True, border=0, new_x=XPos.LMARGIN, new_y=YPos.NEXT)

    annexures = create_annexure_list(data['development'])

    for annexure in annexures:
        pdf.cell(80, 7, f"{annexure['annexure']}", align="L", markdown=True, border=0)
        pdf.cell(110, 7, f"{annexure['description']}", align="L", markdown=True, border=0, new_x=XPos.LMARGIN,
                 new_y=YPos.NEXT)

    pdf.add_page()
    new_right_margin = 100  # Adjust this value as needed
    pdf.set_right_margin(new_right_margin)
    pdf.cell(0, 7, f"**STANDARD TERMS AND CONDITIONS**", align="C", markdown=True, border=0, new_x=XPos.LMARGIN,
             new_y=YPos.NEXT)
    pdf.cell(0, 7, f"", align="C", markdown=True, border=0, new_x=XPos.LMARGIN,
             new_y=YPos.NEXT)

    def initials():
        pdf.set_font('helvetica', '', 8)
        pdf.set_text_color(211, 211, 211)
        pdf.set_line_width(0.5)
        pdf.set_draw_color(211, 211, 211)
        x = pdf.get_x() + 156
        y = pdf.get_y() - 1
        pdf.rect(x, y, 30, 8)
        pdf.cell(0, 7, f"INITIAL                 ", align="R", new_x=XPos.LMARGIN, new_y=YPos.NEXT)
        pdf.set_text_color(0, 0, 0)
        pdf.set_font('helvetica', '', 9)

    standard_conditions = create_standard_conditions(data)

    for item in standard_conditions:
        if item['type'] == 1 and item['section'] != "":
            if pdf.get_y() > 250:
                pdf.add_page()
            pdf.cell(0, 4, f"", align="C", markdown=True, border=0, new_x=XPos.LMARGIN, new_y=YPos.NEXT)
            pdf.cell(10, 7, f"{item['section']}", align="C", markdown=True, border=0)
            pdf.cell(40, 7, f"{item['text']}", align="L", markdown=True, border=0,
                     new_x=XPos.LMARGIN,
                     new_y=YPos.NEXT)

            if item['initial']:
                initials()

        elif item['type'] == 1 and item['section'] == "":
            if pdf.get_y() > 250:
                pdf.add_page()
            pdf.cell(0, 4, f"", align="C", markdown=True, border=0, new_x=XPos.LMARGIN, new_y=YPos.NEXT)
            pdf.cell(10, 7, f"{item['section']}", align="C", markdown=True, border=0)
            pdf.multi_cell(150, 5, f"{item['text']}", align="J", markdown=True, border=0,
                           new_x=XPos.LMARGIN,
                           new_y=YPos.NEXT)
            if item['initial']:
                initials()

        elif item['type'] == 2:
            if pdf.get_y() > 250:
                pdf.add_page()
            pdf.cell(0, 2, f"", align="C", markdown=True, border=0, new_x=XPos.LMARGIN, new_y=YPos.NEXT)
            pdf.cell(10, 7, f"", align="C", markdown=True, border=0)
            pdf.cell(10, 7, f"{item['section']}", align="C", markdown=True, border=0)
            pdf.multi_cell(140, 5,
                           f"{item['text']}",
                           align="J", markdown=True, new_x=XPos.LMARGIN, new_y=YPos.NEXT, border=0)
            if item['initial']:
                initials()
        elif item['type'] == 3:
            if pdf.get_y() > 250:
                pdf.add_page()
            pdf.cell(0, 2, f"", align="C", markdown=True, border=0, new_x=XPos.LMARGIN, new_y=YPos.NEXT)
            pdf.cell(10, 7, f"", align="C", markdown=True, border=0)
            pdf.cell(10, 7, f"", align="C", markdown=True, border=0)
            pdf.cell(10, 7, f"{item['section']}", align="C", markdown=True, border=0)
            pdf.multi_cell(120, 5,
                           f"{item['text']}",
                           align="J", markdown=True, new_x=XPos.LMARGIN, new_y=YPos.NEXT,
                           border=0)

            if item['initial']:
                initials()

    pdf.output(f"sales_client_onboarding_docs/{data['opportunity_code']}-OTP.pdf")

    return f"sales_client_onboarding_docs/{data['opportunity_code']}-OTP.pdf"
