import copy
import os

from fpdf import FPDF, XPos, YPos


class PDF(FPDF):

    def __init__(self, data):
        super().__init__()
        self.data = data
        self.unit = 'mm'
        self.format = 'A4'

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

        self.rect(165, 282, 30, 10)

        # put a bold line above the footer
        self.set_draw_color(0, 0, 0)
        self.set_line_width(0.5)
        self.line(5, self.get_y() - 2, 200, self.get_y() - 2)

    def cell_with_padding(self, w, h, txt, border=0, ln=0, align='', fill=False):
        # Add padding
        padding_left = 0  # Set your desired left padding value
        padding_right = 0  # Set your desired right padding value
        padding_top = 4  # Set your desired top padding value
        padding_bottom = 0  # Set your desired bottom padding value

        # Adjust cell size to include padding
        w += padding_left + padding_right
        h += padding_top + padding_bottom

        self.cell(w, h, txt, border, ln, align, fill)


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
    elif data['development'] == 'Endulini':
        company = 'ENDULINI PROJECTS PROPRIETARY LIMITED'
        company_reg_no = 'REGISTRATION NUMBER 2020/495057/07'
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

    ####################################
    purchaser_signing_details = []
    if data['opportunity_client_type'] == 'Company':
        insert = {
            "purchaser": "Purchaser",
            "marital_status": "N/A",
        }
        purchaser_signing_details.append(insert)
    else:
        if number_of_purchasers >= 1:
            insert = {
                "purchaser": "Purchaser",
                "marital_status": data['opportunity_martial_status'],
                "info": "1ST PURCHASER'S SPOUSE"
            }
            purchaser_signing_details.append(insert)
        if number_of_purchasers >= 2:
            insert = {
                "purchaser": "2nd Purchaser",
                "marital_status": data['opportunity_martial_status_sec'],
                "info": "2ND PURCHASER'S SPOUSE"
            }
            purchaser_signing_details.append(insert)
        if number_of_purchasers >= 3:
            insert = {
                "purchaser": "3rd Purchaser",
                "marital_status": data['opportunity_martial_status_3rd'],
                "info": "3RD PURCHASER'S SPOUSE"
            }
            purchaser_signing_details.append(insert)
        if number_of_purchasers >= 4:
            insert = {
                "purchaser": "4th Purchaser",
                "marital_status": data['opportunity_martial_status_4th'],
                "info": "4TH PURCHASER'S SPOUSE"
            }
            purchaser_signing_details.append(insert)
        if number_of_purchasers >= 5:
            insert = {
                "purchaser": "5th Purchaser",
                "marital_status": data['opportunity_martial_status_5th'],
                "info": "5TH PURCHASER'S SPOUSE"
            }
            purchaser_signing_details.append(insert)
        if number_of_purchasers >= 6:
            insert = {
                "purchaser": "6th Purchaser",
                "marital_status": data['opportunity_martial_status_6th'],
                "info": "6TH PURCHASER'S SPOUSE"
            }
            purchaser_signing_details.append(insert)
        if number_of_purchasers >= 7:
            insert = {
                "purchaser": "7th Purchaser",
                "marital_status": data['opportunity_martial_status_7th'],
                "info": "7TH PURCHASER'S SPOUSE"
            }
            purchaser_signing_details.append(insert)
        if number_of_purchasers >= 8:
            insert = {
                "purchaser": "8th Purchaser",
                "marital_status": data['opportunity_martial_status_8th'],
                "info": "8TH PURCHASER'S SPOUSE"
            }
            purchaser_signing_details.append(insert)
        if number_of_purchasers >= 9:
            insert = {
                "purchaser": "9th Purchaser",
                "marital_status": data['opportunity_martial_status_9th'],
                "info": "9TH PURCHASER'S SPOUSE"
            }
            purchaser_signing_details.append(insert)
        if number_of_purchasers >= 10:
            insert = {
                "purchaser": "10th Purchaser",
                "marital_status": data['opportunity_martial_status_10th'],
                "info": "10TH PURCHASER'S SPOUSE"
            }
            purchaser_signing_details.append(insert)

    ####################################

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
                   align="L", markdown=True,
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
                               f"**{data['opportunity_company_postal_address']}**", align="L", markdown=True,
                       new_x=XPos.LMARGIN, new_y=YPos.NEXT, border=1)

        pdf.cell(20, 7, "B4", align="C", markdown=True, border=1)
        pdf.cell(50, 7, "Telehpone", align="L", markdown=True, border=1)
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
        purchaser_details = []
        if number_of_purchasers >= 1:
            insert = {
                "header": "Purchaser",
                "label_b1": "Full Name",
                "value_b1": f"{data['opportunity_firstname']} {data['opportunity_lastname']}",
                "label_b2": "ID Number",
                "value_b2": data['opportunity_id'],
                "label_b3": "Address (Street & Postal)",
                "value_b3": f"{data['opportunity_residental_address']} & {data['opportunity_postal_address']}",
                "label_b4": "Marital Status",
                "value_b4": data['opportunity_martial_status'],
                "label_b5": "Telephone",
                "value_b5": data['opportunity_landline'],
                "label_b6": "Mobile",
                "value_b6": data['opportunity_mobile'],
                "label_b7": "E-mail",
                "value_b7": data['opportunity_email'],
            }
            purchaser_details.append(insert)

        if number_of_purchasers >= 2:
            insert = {
                "header": "2nd Purchaser",
                "label_b1": "Full Name",
                "value_b1": f"{data['opportunity_firstname_sec']} {data['opportunity_lastname_sec']}",
                "label_b2": "ID Number",
                "value_b2": data['opportunity_id_sec'],
                "label_b3": "Address (Street & Postal)",
                "value_b3": f"{data['opportunity_residental_address_sec']} & {data['opportunity_postal_address_sec']}",
                "label_b4": "Marital Status",
                "value_b4": data['opportunity_martial_status_sec'],
                "label_b5": "Telephone",
                "value_b5": data['opportunity_landline_sec'],
                "label_b6": "Mobile",
                "value_b6": data['opportunity_mobile_sec'],
                "label_b7": "E-mail",
                "value_b7": data['opportunity_email_sec'],
            }
            purchaser_details.append(insert)

        if number_of_purchasers >= 3:
            insert = {
                "header": "3rd Purchaser",
                "label_b1": "Full Name",
                "value_b1": f"{data['opportunity_firstname_3rd']} {data['opportunity_lastname_3rd']}",
                "label_b2": "ID Number",
                "value_b2": data['opportunity_id_3rd'],
                "label_b3": "Address (Street & Postal)",
                "value_b3": f"{data['opportunity_residental_address_3rd']} & {data['opportunity_postal_address_3rd']}",
                "label_b4": "Marital Status",
                "value_b4": data['opportunity_martial_status_3rd'],
                "label_b5": "Telephone",
                "value_b5": data['opportunity_landline_3rd'],
                "label_b6": "Mobile",
                "value_b6": data['opportunity_mobile_3rd'],
                "label_b7": "E-mail",
                "value_b7": data['opportunity_email_3rd'],
            }
            purchaser_details.append(insert)

        if number_of_purchasers >= 4:
            insert = {
                "header": "4th Purchaser",
                "label_b1": "Full Name",
                "value_b1": f"{data['opportunity_firstname_4th']} {data['opportunity_lastname_4th']}",
                "label_b2": "ID Number",
                "value_b2": data['opportunity_id_4th'],
                "label_b3": "Address (Street & Postal)",
                "value_b3": f"{data['opportunity_residental_address_4th']} & {data['opportunity_postal_address_4th']}",
                "label_b4": "Marital Status",
                "value_b4": data['opportunity_martial_status_4th'],
                "label_b5": "Telephone",
                "value_b5": data['opportunity_landline_4th'],
                "label_b6": "Mobile",
                "value_b6": data['opportunity_mobile_4th'],
                "label_b7": "E-mail",
                "value_b7": data['opportunity_email_4th'],
            }
            purchaser_details.append(insert)

        if number_of_purchasers >= 5:
            insert = {
                "header": "5th Purchaser",
                "label_b1": "Full Name",
                "value_b1": f"{data['opportunity_firstname_5th']} {data['opportunity_lastname_5th']}",
                "label_b2": "ID Number",
                "value_b2": data['opportunity_id_5th'],
                "label_b3": "Address (Street & Postal)",
                "value_b3": f"{data['opportunity_residental_address_5th']} & {data['opportunity_postal_address_5th']}",
                "label_b4": "Marital Status",
                "value_b4": data['opportunity_martial_status_5th'],
                "label_b5": "Telephone",
                "value_b5": data['opportunity_landline_5th'],
                "label_b6": "Mobile",
                "value_b6": data['opportunity_mobile_5th'],
                "label_b7": "E-mail",
                "value_b7": data['opportunity_email_5th'],
            }
            purchaser_details.append(insert)

        if number_of_purchasers >= 6:
            insert = {
                "header": "6th Purchaser",
                "label_b1": "Full Name",
                "value_b1": f"{data['opportunity_firstname_6th']} {data['opportunity_lastname_6th']}",
                "label_b2": "ID Number",
                "value_b2": data['opportunity_id_6th'],
                "label_b3": "Address (Street & Postal)",
                "value_b3": f"{data['opportunity_residental_address_6th']} & {data['opportunity_postal_address_6th']}",
                "label_b4": "Marital Status",
                "value_b4": data['opportunity_martial_status_6th'],
                "label_b5": "Telephone",
                "value_b5": data['opportunity_landline_6th'],
                "label_b6": "Mobile",
                "value_b6": data['opportunity_mobile_6th'],
                "label_b7": "E-mail",
                "value_b7": data['opportunity_email_6th'],
            }
            purchaser_details.append(insert)

        if number_of_purchasers >= 7:
            insert = {
                "header": "7th Purchaser",
                "label_b1": "Full Name",
                "value_b1": f"{data['opportunity_firstname_7th']} {data['opportunity_lastname_7th']}",
                "label_b2": "ID Number",
                "value_b2": data['opportunity_id_7th'],
                "label_b3": "Address (Street & Postal)",
                "value_b3": f"{data['opportunity_residental_address_7th']} & {data['opportunity_postal_address_7th']}",
                "label_b4": "Marital Status",
                "value_b4": data['opportunity_martial_status_7th'],
                "label_b5": "Telephone",
                "value_b5": data['opportunity_landline_7th'],
                "label_b6": "Mobile",
                "value_b6": data['opportunity_mobile_7th'],
                "label_b7": "E-mail",
                "value_b7": data['opportunity_email_7th'],
            }
            purchaser_details.append(insert)

        if number_of_purchasers >= 8:
            insert = {
                "header": "8th Purchaser",
                "label_b1": "Full Name",
                "value_b1": f"{data['opportunity_firstname_8th']} {data['opportunity_lastname_8th']}",
                "label_b2": "ID Number",
                "value_b2": data['opportunity_id_8th'],
                "label_b3": "Address (Street & Postal)",
                "value_b3": f"{data['opportunity_residental_address_8th']} & {data['opportunity_postal_address_8th']}",
                "label_b4": "Marital Status",
                "value_b4": data['opportunity_martial_status_8th'],
                "label_b5": "Telephone",
                "value_b5": data['opportunity_landline_8th'],
                "label_b6": "Mobile",
                "value_b6": data['opportunity_mobile_8th'],
                "label_b7": "E-mail",
                "value_b7": data['opportunity_email_8th'],
            }
            purchaser_details.append(insert)

        if number_of_purchasers >= 9:
            insert = {
                "header": "9th Purchaser",
                "label_b1": "Full Name",
                "value_b1": f"{data['opportunity_firstname_9th']} {data['opportunity_lastname_9th']}",
                "label_b2": "ID Number",
                "value_b2": data['opportunity_id_9th'],
                "label_b3": "Address (Street & Postal)",
                "value_b3": f"{data['opportunity_residental_address_9th']} & {data['opportunity_postal_address_9th']}",
                "label_b4": "Marital Status",
                "value_b4": data['opportunity_martial_status_9th'],
                "label_b5": "Telephone",
                "value_b5": data['opportunity_landline_9th'],
                "label_b6": "Mobile",
                "value_b6": data['opportunity_mobile_9th'],
                "label_b7": "E-mail",
                "value_b7": data['opportunity_email_9th'],
            }
            purchaser_details.append(insert)

        if number_of_purchasers >= 10:
            insert = {
                "header": "10th Purchaser",
                "label_b1": "Full Name",
                "value_b1": f"{data['opportunity_firstname_10th']} {data['opportunity_lastname_10th']}",
                "label_b2": "ID Number",
                "value_b2": data['opportunity_id_10th'],
                "label_b3": "Address (Street & Postal)",
                "value_b3": f"{data['opportunity_residental_address_10th']} & {data['opportunity_postal_address_10th']}",
                "label_b4": "Marital Status",
                "value_b4": data['opportunity_martial_status_10th'],
                "label_b5": "Telephone",
                "value_b5": data['opportunity_landline_10th'],
                "label_b6": "Mobile",
                "value_b6": data['opportunity_mobile_10th'],
                "label_b7": "E-mail",
                "value_b7": data['opportunity_email_10th'],
            }
            purchaser_details.append(insert)

        ##########################################

        # Add the purchaser details to the pdf
        for purchaser in purchaser_details:
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
            pdf.multi_cell(120, 5, f"**{purchaser['value_b3']}**", align="L", markdown=True,
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
            pdf.cell(120, 7, f"**{purchaser['value_b7']}**", align="L", markdown=True,
                     new_x=XPos.LMARGIN, new_y=YPos.NEXT, border=1)

            pdf.cell(0, 5, "", new_x=XPos.LMARGIN, new_y=YPos.NEXT)

    pdf.add_page()

    pdf.cell(20, 7, f"**C**", align="C", markdown=True, border=1)
    pdf.cell(50, 7, f"**The Property**", align="L", markdown=True, border=1)
    pdf.cell(120, 7, "", align="C", markdown=True, new_x=XPos.LMARGIN, new_y=YPos.NEXT, border=1)

    pdf.cell(20, 15, f"C1", align="C", markdown=True, border=1)
    pdf.cell(50, 15, f"Unit", align="L", markdown=True, border=1)
    pdf.multi_cell(120, 5,
                   f"**Unit No:**{door_no} having an approximate floor area of .......... Square Metre's as "
                   f"reflected on the Development and Unit Plans annexed hereto (marked \"**A & B**\")",
                   align="C", markdown=True, new_x=XPos.LMARGIN, new_y=YPos.NEXT, border=1)

    # if data['opportunity_gardenNumber'] != "":

    pdf.output(f"sales_client_onboarding_docs/{data['opportunity_code']}-OTP.pdf")

    return f"sales_client_onboarding_docs/{data['opportunity_code']}-OTP.pdf"
