# import copy
# import os

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
                "value_b3": f"{data['opportunity_residental_address_10th']} & "
                            f"{data['opportunity_postal_address_10th']}",
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
                   align="C", markdown=True, new_x=XPos.LMARGIN, new_y=YPos.NEXT, border=1)

    ##############################################

    garden_number = 0
    garden_size = 0

    if data['opportunity_gardenNumber'] != 0 and data['opportunity_gardenNumber'] is not None:
        garden_number = data['opportunity_gardenNumber']
        garden_size = data['opportunity_gardenSize']

    aditional_parking_bays = data['opportunity_additional_bay']
    allocated_parking_bay = data['opportunity_originalBayNo']
    second_parking_bay = data['opportunity_parking_base']
    third_parking_bay = data['opportunity_parking_base2']

    line_item = 1

    if garden_number != "0":
        pdf.cell(20, 20, f"C2", align="C", markdown=True, border="LTR")
        pdf.cell(50, 20, f"Exclusive Use Area's", align="L", markdown=True, border="LTR")
        pdf.multi_cell(120, 5,
                       f"**{line_item}. Garden G {garden_number}** having an area of approximately {garden_size} "
                       f"square metres and shall be allocated to the Purchaser in terms of Section 27A of the "
                       f"Sectional Titles Act.  The Garden is indicated on the Exclusive Use Area Plan annexed hereto "
                       f"(marked **\"A\")**.",
                       align="C", markdown=True, new_x=XPos.LMARGIN, new_y=YPos.NEXT, border="LTR")

        line_item += 1
        if aditional_parking_bays != 0:
            border = "LR"
        else:
            border = "LBR"

        pdf.cell(20, 20, f"", align="C", markdown=True, border=border)
        pdf.cell(50, 20, f"", align="L", markdown=True, border=border)
        pdf.multi_cell(120, 5,
                       f"**{line_item}. Parking P {allocated_parking_bay}** having an area of approximately .........."
                       f".. square metres and shall be allocated to the Purchaser in terms of Section 27A of the "
                       f"Sectional Titles Act.  The Parking is indicated on the Exclusive Use Area Plan annexed "
                       f"hereto (marked **\"A\"**).",
                       align="C", markdown=True, new_x=XPos.LMARGIN, new_y=YPos.NEXT, border=border)

        line_item += 1
        if aditional_parking_bays == 1:
            border = "LBR"

        elif aditional_parking_bays == 2:
            border = "LR"

        if aditional_parking_bays >= 1:
            pdf.cell(20, 20, f"", align="C", markdown=True, border=border)
            pdf.cell(50, 20, f"", align="L", markdown=True, border=border)
            pdf.multi_cell(120, 5,
                           f"**{line_item}. Additional Parking P {second_parking_bay}** having an area of approximately"
                           f" ............ square metres and shall be allocated to the Purchaser in terms of Section "
                           f"27A of the Sectional Titles Act.  The Parking is indicated on the Exclusive Use Area "
                           f"Plan annexed hereto (marked **\"A\"**).",
                           align="C", markdown=True, new_x=XPos.LMARGIN, new_y=YPos.NEXT, border=border)

        if aditional_parking_bays == 2:
            line_item += 1
            border = "LBR"
            pdf.cell(20, 20, f"", align="C", markdown=True, border=border)
            pdf.cell(50, 20, f"", align="L", markdown=True, border=border)
            pdf.multi_cell(120, 5,
                           f"**{line_item}. Additional Parking P {third_parking_bay}** having an area of approximately "
                           f"............ square metres and shall be allocated to the Purchaser in terms of Section "
                           f"27A of the Sectional Titles Act.  The Parking is indicated on the Exclusive Use Area "
                           f"Plan annexed hereto (marked **\"A\"**).",
                           align="C", markdown=True, new_x=XPos.LMARGIN, new_y=YPos.NEXT, border=border)

    else:
        line_item = 1
        if aditional_parking_bays != 0:
            border = "LR"
        else:
            border = "LBR"

        pdf.cell(20, 20, f"C2", align="C", markdown=True, border=border)
        pdf.cell(50, 20, f"Exclusive Use Area's", align="L", markdown=True, border=border)
        pdf.multi_cell(120, 5,
                       f"**{line_item}. Parking P {allocated_parking_bay}** having an area of approximately .........."
                       f".. square metres and shall be allocated to the Purchaser in terms of Section 27A of the "
                       f"Sectional Titles Act.  The Parking is indicated on the Exclusive Use Area Plan annexed "
                       f"hereto (marked **\"A\"**).",
                       align="C", markdown=True, new_x=XPos.LMARGIN, new_y=YPos.NEXT, border=border)

        line_item += 1
        if aditional_parking_bays == 1:
            border = "LBR"

        elif aditional_parking_bays == 2:
            border = "LR"

        if aditional_parking_bays >= 1:
            pdf.cell(20, 20, f"", align="C", markdown=True, border=border)
            pdf.cell(50, 20, f"", align="L", markdown=True, border=border)
            pdf.multi_cell(120, 5,
                           f"**{line_item}. Additional Parking P {second_parking_bay}** having an area of approximately "
                           f"............ square metres and shall be allocated to the Purchaser in terms of Section "
                           f"27A of the Sectional Titles Act.  The Parking is indicated on the Exclusive Use Area "
                           f"Plan annexed hereto (marked **\"A\"**).",
                           align="C", markdown=True, new_x=XPos.LMARGIN, new_y=YPos.NEXT, border=border)

        if aditional_parking_bays == 2:
            line_item += 1
            border = "LBR"
            pdf.cell(20, 20, f"", align="C", markdown=True, border=border)
            pdf.cell(50, 20, f"", align="L", markdown=True, border=border)
            pdf.multi_cell(120, 5,
                           f"**{line_item}. Additional Parking P {third_parking_bay}** having an area of approximately "
                           f"............ square metres and shall be allocated to the Purchaser in terms of Section "
                           f"27A of the Sectional Titles Act.  The Parking is indicated on the Exclusive Use Area "
                           f"Plan annexed hereto (marked **\"A\"**).",
                           align="C", markdown=True, new_x=XPos.LMARGIN, new_y=YPos.NEXT, border=border)

    ##############################################

    pdf.cell(20, 7, f"C3", align="C", markdown=True, border=1)
    pdf.cell(50, 7, f"Street Address", align="L", markdown=True, border=1)
    pdf.multi_cell(120, 7,
                   f"**Unit No.** {door_no}, {data['development']}, {street_address}",
                   align="C", markdown=True, new_x=XPos.LMARGIN, new_y=YPos.NEXT, border=1)

    pdf.cell(0, 3, f"", align="C", markdown=True, border=0, new_x=XPos.LMARGIN, new_y=YPos.NEXT)

    pdf.cell(20, 8, f"**D**", align="C", markdown=True, border=1)
    pdf.cell(50, 8, f"**Estimated Occupation**", align="L", markdown=True, border=1)
    pdf.multi_cell(120, 4,
                   f"**As per and subject to Clause 11.13 of the Agreement Estimated to be on or around:............",
                   align="L", markdown=True, new_x=XPos.LMARGIN, new_y=YPos.NEXT, border=1)

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

    pdf.cell(20, 15, "**G**", align="C", markdown=True, border=1)
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

    pdf.cell(20, 7, f"**G**", align="C", markdown=True, border=1)
    pdf.cell(50, 7, f"**Estimated Monthly Levy**", align="L", markdown=True, border=1)
    pdf.cell(120, 7, f"**TBC By Managing Agent (Payable by Purchaser)** Approx {estimated_levy}", align="L",
             markdown=True, new_x=XPos.LMARGIN, new_y=YPos.NEXT, border=1)

    pdf.cell(0, 3, f"", align="C", markdown=True, border=0, new_x=XPos.LMARGIN, new_y=YPos.NEXT)

    pdf.cell(20, 7, f"**H**", align="C", markdown=True, border=1)
    pdf.cell(50, 7, f"**Electricity Deposit**", align="L", markdown=True, border=1)
    pdf.cell(120, 7, f"**R1200.00 Payable upon Registration to the Transfer Attorneys", align="L",
             markdown=True, new_x=XPos.LMARGIN, new_y=YPos.NEXT, border=1)

    pdf.cell(0, 3, f"", align="C", markdown=True, border=0, new_x=XPos.LMARGIN, new_y=YPos.NEXT)

    pdf.cell(20, 7, f"**I**", align="C", markdown=True, border=1)
    pdf.cell(50, 7, f"**Mortgage Originator**", align="L", markdown=True, border=1)
    pdf.cell(120, 7, f"{data['opportunity_bond_originator']}", align="L",
             markdown=True, new_x=XPos.LMARGIN, new_y=YPos.NEXT, border=1)

    # extras
    print(data['opportunity_extras_not_listed'][1:])
    # notes
    print(data['opportunity_notes'])
    # Gas
    print(data['opportunity_stove_option'])
    # specials
    print(data['opportunity_specials'])

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
    pdf.cell(80, 7, f"**Annexure A**", align="L", markdown=True, border=0)
    pdf.cell(110, 7, f"Site Layout Plan", align="L", markdown=True, border=0, new_x=XPos.LMARGIN, new_y=YPos.NEXT)
    pdf.cell(80, 7, f"**Annexure B**", align="L", markdown=True, border=0)
    pdf.cell(110, 7, f"SDP & Parking Correlation", align="L", markdown=True, border=0,
             new_x=XPos.LMARGIN, new_y=YPos.NEXT)
    pdf.cell(80, 7, f"**Annexure C**", align="L", markdown=True, border=0)
    pdf.cell(110, 7, f"Unit Plan", align="L", markdown=True, border=0, new_x=XPos.LMARGIN, new_y=YPos.NEXT)
    pdf.cell(80, 7, f"**Annexure D**", align="L", markdown=True, border=0)
    pdf.cell(110, 7, f"Finishes Schedule", align="L", markdown=True, border=0, new_x=XPos.LMARGIN, new_y=YPos.NEXT)
    pdf.cell(80, 7, f"**Annexure E**", align="L", markdown=True, border=0)
    pdf.cell(110, 7, f"Specifications of Finishes", align="L", markdown=True, border=0,
             new_x=XPos.LMARGIN, new_y=YPos.NEXT)
    pdf.cell(80, 7, f"**Annexure F**", align="L", markdown=True, border=0)
    pdf.cell(110, 7, f"Consent in terms of the Protection of Personal Information Act",
             align="L", markdown=True, border=0, new_x=XPos.LMARGIN, new_y=YPos.NEXT)

    pdf.add_page()

    pdf.cell(0, 7, f"**STANDARD TERMS AND CONDITIONS**", align="C", markdown=True, border=0, new_x=XPos.LMARGIN, new_y=YPos.NEXT)
    pdf.cell(0, 7, f"", align="C", markdown=True, border=0, new_x=XPos.LMARGIN,
             new_y=YPos.NEXT)

    pdf.cell(20, 7, f"1.", align="C", markdown=True, border=0)
    pdf.cell(50, 7, f"**PREAMBLE**", align="L", markdown=True, border=0, new_x=XPos.LMARGIN,
             new_y=YPos.NEXT)

    pdf.cell(20, 7, f"", align="C", markdown=True, border=0)
    pdf.cell(20, 7, f"1.1", align="C", markdown=True, border=0)
    pdf.multi_cell(150, 5, f"The Seller has agreed to sell, and the Purchaser has agreed to purchase the Property in C"
                           f" of the Information Schedule, to be established in the Sectional Title Scheme to be "
                           f"known as **\"{data['development'].upper()}\"** in terms of the Sectional Titles Act.",
                     align="J", markdown=True, new_x=XPos.LMARGIN, new_y=YPos.NEXT, border=0)

    pdf.cell(20, 7, f"", align="C", markdown=True, border=0)
    pdf.cell(20, 7, f"1.2", align="C", markdown=True, border=0)
    pdf.multi_cell(150, 5, f"The Sale is subject to the fulfilment of the condition's precedent recorded "
                           f"in this Agreement.",
                   align="J", markdown=True, new_x=XPos.LMARGIN, new_y=YPos.NEXT, border=0)

    pdf.cell(20, 7, f"2.", align="C", markdown=True, border=0)
    pdf.cell(50, 7, f"**INTERPRETATION**", align="L", markdown=True, border=0, new_x=XPos.LMARGIN,
             new_y=YPos.NEXT)

    pdf.cell(20, 7, f"", align="C", markdown=True, border=0)
    pdf.cell(20, 7, f"2.1", align="C", markdown=True, border=0)
    pdf.multi_cell(150, 5, f"In this Agreement, unless the context otherwise indicates:",
                   align="J", markdown=True, new_x=XPos.LMARGIN, new_y=YPos.NEXT, border=0)

    pdf.cell(20, 7, f"", align="C", markdown=True, border=0)
    pdf.cell(20, 7, f"", align="C", markdown=True, border=0)
    pdf.cell(20, 7, f"2.1.1", align="C", markdown=True, border=0)
    pdf.multi_cell(130, 5, f"**\"Architect\"** means any registered architect as may be appointed by the Seller "
                           f"from time to time.", align="J", markdown=True, new_x=XPos.LMARGIN, new_y=YPos.NEXT,
                   border=0)

    pdf.cell(20, 7, f"", align="C", markdown=True, border=0)
    pdf.cell(20, 7, f"", align="C", markdown=True, border=0)
    pdf.cell(20, 7, f"2.1.2", align="C", markdown=True, border=0)
    pdf.multi_cell(130, 5, f"**\"Beneficial Occupation\"** means the Property has water, power, sewerage, access and "
                           f"is thus liveable and ready for physical occupation.",
                   align="J", markdown=True, new_x=XPos.LMARGIN, new_y=YPos.NEXT,
                   border=0)

    pdf.cell(20, 7, f"", align="C", markdown=True, border=0)
    pdf.cell(20, 7, f"", align="C", markdown=True, border=0)
    pdf.cell(20, 7, f"2.1.3", align="C", markdown=True, border=0)
    pdf.multi_cell(130, 5, f"**\"Body Corporate\"** means the controlling body of the Scheme as contemplated in terms"
                           f" of Section 36 of the Sectional Titles Act, which will come into existence with the"
                           f" transfer of the first Unit from the Seller to a Purchaser in this Scheme.",
                   align="J", markdown=True, new_x=XPos.LMARGIN, new_y=YPos.NEXT,
                   border=0)

    pdf.cell(20, 7, f"", align="C", markdown=True, border=0)
    pdf.cell(20, 7, f"", align="C", markdown=True, border=0)
    pdf.cell(20, 7, f"2.1.4", align="C", markdown=True, border=0)
    pdf.multi_cell(130, 5, f"**\"Building\"** means the building/s to be constructed on the Land.",
                   align="J", markdown=True, new_x=XPos.LMARGIN, new_y=YPos.NEXT,
                   border=0)

    pdf.cell(20, 7, f"", align="C", markdown=True, border=0)
    pdf.cell(20, 7, f"", align="C", markdown=True, border=0)
    pdf.cell(20, 7, f"2.1.5", align="C", markdown=True, border=0)
    pdf.multi_cell(130, 5, f"**\"Chief Ombud\"** means Chief Ombud as defined in Section 1 of the Community Schemes Ombud Service Act, 2010.",
                   align="J", markdown=True, new_x=XPos.LMARGIN, new_y=YPos.NEXT,
                   border=0)

    pdf.cell(20, 7, f"", align="C", markdown=True, border=0)
    pdf.cell(20, 7, f"", align="C", markdown=True, border=0)
    pdf.cell(20, 7, f"2.1.6", align="C", markdown=True, border=0)
    pdf.multi_cell(130, 5,
                   f"**\"Completion Date\"** means the date upon which the building inspector employed by the local authority or Architect issues an Occupation Certificate in respect of the Unit to the effect that the Unit is fit for Beneficial Occupation, or the date of handover of the keys of the Unit to the Purchaser, whichever date is earlier, subject to the provision that in the event of a dispute, the Completion Date shall be certified as such by the Architect, whose decision as to that date shall be final and binding on the parties.",
                   align="J", markdown=True, new_x=XPos.LMARGIN, new_y=YPos.NEXT,
                   border=0)

    pdf.cell(20, 7, f"", align="C", markdown=True, border=0)
    pdf.cell(20, 7, f"", align="C", markdown=True, border=0)
    pdf.cell(20, 7, f"2.1.7", align="C", markdown=True, border=0)
    pdf.multi_cell(130, 5,
                   f"**\"Common Property\"** means common property as defined in the Sectional Titles Act.",
                   align="J", markdown=True, new_x=XPos.LMARGIN, new_y=YPos.NEXT,
                   border=0)

    pdf.cell(20, 7, f"", align="C", markdown=True, border=0)
    pdf.cell(20, 7, f"", align="C", markdown=True, border=0)
    pdf.cell(20, 7, f"2.1.8", align="C", markdown=True, border=0)
    pdf.multi_cell(130, 5,
                   f"**\"Developer\"** means the party described as Seller in the Information Schedule.",
                   align="J", markdown=True, new_x=XPos.LMARGIN, new_y=YPos.NEXT,
                   border=0)

    pdf.cell(20, 7, f"", align="C", markdown=True, border=0)
    pdf.cell(20, 7, f"", align="C", markdown=True, border=0)
    pdf.cell(20, 7, f"2.1.9", align="C", markdown=True, border=0)
    pdf.multi_cell(130, 5,
                   f"**\"Development Period\"** means the period from the establishment of the Body Corporate to the transfer of the last saleable sectional title unit in the Scheme or a period not exceeding twenty years from date of establishment of the Body Corporate, whichever is the longest.",
                   align="J", markdown=True, new_x=XPos.LMARGIN, new_y=YPos.NEXT,
                   border=0)

    pdf.cell(20, 7, f"", align="C", markdown=True, border=0)
    pdf.cell(20, 7, f"", align="C", markdown=True, border=0)
    pdf.cell(20, 7, f"2.1.10", align="C", markdown=True, border=0)
    pdf.multi_cell(130, 5,
                   f"**\"Exclusive Use Areas\"** means such parts of the Common Property reserved for the exclusive use and enjoyment of the registered owner for the time being of the Unit, in terms of Section 27A of the Sectional Titles Act which includes the Garden and Parking.",
                   align="J", markdown=True, new_x=XPos.LMARGIN, new_y=YPos.NEXT,
                   border=0)

    pdf.cell(20, 7, f"", align="C", markdown=True, border=0)
    pdf.cell(20, 7, f"", align="C", markdown=True, border=0)
    pdf.cell(20, 7, f"2.1.11", align="C", markdown=True, border=0)
    pdf.multi_cell(130, 5,
                   f"**\"FICA\"** means the Financial Intelligence Centre Act, Act 38 of 2001, as amended from time to time.",
                   align="J", markdown=True, new_x=XPos.LMARGIN, new_y=YPos.NEXT,
                   border=0)

    pdf.cell(20, 7, f"", align="C", markdown=True, border=0)
    pdf.cell(20, 7, f"", align="C", markdown=True, border=0)
    pdf.cell(20, 7, f"2.1.12", align="C", markdown=True, border=0)
    pdf.multi_cell(130, 5,
                   f"**\"Happy Letter\"** means a formal document prepared in a format acceptable to the bank or other recognised financial institution providing a bond to the Purchaser as provided for in Clause 19 hereunder and signed by the Purchaser (or his Agent/Proxy) at the instance of the bank or other recognised financial institution providing a bond, or in the case of a cash purchase, a Happy Letter document provided by the Seller, to be signed by the Purchaser.",
                   align="J", markdown=True, new_x=XPos.LMARGIN, new_y=YPos.NEXT,
                   border=0)

    pdf.cell(20, 7, f"", align="C", markdown=True, border=0)
    pdf.cell(20, 7, f"", align="C", markdown=True, border=0)
    pdf.cell(20, 7, f"2.1.13", align="C", markdown=True, border=0)
    pdf.multi_cell(130, 5,
                   f"**\"Information Schedule\"** means the Information Schedule set out on pages 3, 4 and 5 hereof which shall be deemed to be incorporated in this Agreement and shall be an integral part thereof.",
                   align="J", markdown=True, new_x=XPos.LMARGIN, new_y=YPos.NEXT,
                   border=0)

    pdf.cell(20, 7, f"", align="C", markdown=True, border=0)
    pdf.cell(20, 7, f"", align="C", markdown=True, border=0)
    pdf.cell(20, 7, f"2.1.14", align="C", markdown=True, border=0)
    pdf.multi_cell(130, 5,
                   f"**\"Land\"** means the land on which the Scheme is to be developed being Erf 41409 Kraaifontein, (previously known as the Remainder of Portion 18 of the Farm Langeberg Number 311.",
                   align="J", markdown=True, new_x=XPos.LMARGIN, new_y=YPos.NEXT,
                   border=0)

    pdf.cell(20, 7, f"", align="C", markdown=True, border=0)
    pdf.cell(20, 7, f"", align="C", markdown=True, border=0)
    pdf.cell(20, 7, f"2.1.15", align="C", markdown=True, border=0)
    pdf.multi_cell(130, 5,
                   f"**\"Occupation Certificate\"** means a certificate issued by the City of Cape Town confirming that the Unit is ready for Beneficial Occupation.",
                   align="J", markdown=True, new_x=XPos.LMARGIN, new_y=YPos.NEXT,
                   border=0)

    pdf.cell(20, 7, f"", align="C", markdown=True, border=0)
    pdf.cell(20, 7, f"", align="C", markdown=True, border=0)
    pdf.cell(20, 7, f"2.1.16", align="C", markdown=True, border=0)
    pdf.multi_cell(130, 5,
                   f"**\"Occupation Date\"** means the date on which the Seller hands over the keys of the Unit to the Purchaser, or Transfer Date, whichever is the earliest.",
                   align="J", markdown=True, new_x=XPos.LMARGIN, new_y=YPos.NEXT,
                   border=0)

    pdf.add_page()

    pdf.cell(20, 7, f"", align="C", markdown=True, border=0)
    pdf.cell(20, 7, f"", align="C", markdown=True, border=0)
    pdf.cell(20, 7, f"2.1.17", align="C", markdown=True, border=0)
    pdf.multi_cell(130, 5,
                   f"**\"Prime Rate\"** means a rate of interest per annum which is equal to the published minimum lending rate of interest per annum, compounded monthly in arrear, charged by ABSA Bank Limited on the unsecured overdrawn current accounts of its most favoured corporate clients in the private sector from time to time.  (In the case of a dispute as to the rate so payable, the rate may be certified by any manager or assistant manager of any branch of the said bank, whose decision shall be final and binding on the parties.)",
                   align="J", markdown=True, new_x=XPos.LMARGIN, new_y=YPos.NEXT,
                   border=0)

    pdf.cell(20, 7, f"", align="C", markdown=True, border=0)
    pdf.cell(20, 7, f"", align="C", markdown=True, border=0)
    pdf.cell(20, 7, f"2.1.18", align="C", markdown=True, border=0)
    pdf.multi_cell(130, 5,
                   f"**\"Property\"** means the Property as described in C of the Information Schedule.",
                   align="J", markdown=True, new_x=XPos.LMARGIN, new_y=YPos.NEXT,
                   border=0)

    pdf.cell(20, 7, f"", align="C", markdown=True, border=0)
    pdf.cell(20, 7, f"", align="C", markdown=True, border=0)
    pdf.cell(20, 7, f"2.1.19", align="C", markdown=True, border=0)
    pdf.multi_cell(130, 5,
                   f"**\"Purchaser\"** means the party/ies described as such in B of the Information Schedule.",
                   align="J", markdown=True, new_x=XPos.LMARGIN, new_y=YPos.NEXT,
                   border=0)

    pdf.cell(20, 7, f"", align="C", markdown=True, border=0)
    pdf.cell(20, 7, f"", align="C", markdown=True, border=0)
    pdf.cell(20, 7, f"2.1.20", align="C", markdown=True, border=0)
    pdf.multi_cell(130, 5,
                   f"**\"Rules\"** means the Management and Conduct Rules as Amended for the **\"{data['development'].upper()}\"**  Sectional Title Scheme as prescribed in terms of Section 10(2)(a) and (b) of the Sectional Titles Schemes Management Act No. 8 of 2011, subject to the approval the Chief Ombud, and shall include any substituting rules, available on request from the Agent or Seller.",
                   align="J", markdown=True, new_x=XPos.LMARGIN, new_y=YPos.NEXT,
                   border=0)

    pdf.cell(20, 7, f"", align="C", markdown=True, border=0)
    pdf.cell(20, 7, f"", align="C", markdown=True, border=0)
    pdf.cell(20, 7, f"2.1.21", align="C", markdown=True, border=0)
    pdf.multi_cell(130, 5,
                   f"**\"Scheme\"** means the **\"{data['development'].upper()}\"** Sectional Title Development to be established on the Land, comprising of sectional title residential Units and Exclusive Use Areas, which development may take place in phases and which is situated on the Land as depicted on the Locality, Development, Unit and Exclusive Use Area Plans, Annexure's **\"A,B & C\"** hereto",
                   align="J", markdown=True, new_x=XPos.LMARGIN, new_y=YPos.NEXT,
                   border=0)

    pdf.cell(20, 7, f"", align="C", markdown=True, border=0)
    pdf.cell(20, 7, f"", align="C", markdown=True, border=0)
    pdf.cell(20, 7, f"2.1.22", align="C", markdown=True, border=0)
    pdf.multi_cell(130, 5,
                   f"**\"Sectional Titles Act\"** means the Sectional Titles Act No 95 of 1986 (or any statutory modification or re-enactment thereof) and includes the regulations made thereunder from time to time.",
                   align="J", markdown=True, new_x=XPos.LMARGIN, new_y=YPos.NEXT,
                   border=0)

    pdf.cell(20, 7, f"", align="C", markdown=True, border=0)
    pdf.cell(20, 7, f"", align="C", markdown=True, border=0)
    pdf.cell(20, 7, f"2.1.23", align="C", markdown=True, border=0)
    pdf.multi_cell(130, 5,
                   f"**\"Sectional Plan\"** means the sectional plan/s prepared and registered in respect of the Scheme and includes extension plans to be registered in respect of the Scheme.",
                   align="J", markdown=True, new_x=XPos.LMARGIN, new_y=YPos.NEXT,
                   border=0)

    pdf.cell(20, 7, f"", align="C", markdown=True, border=0)
    pdf.cell(20, 7, f"", align="C", markdown=True, border=0)
    pdf.cell(20, 7, f"2.1.24", align="C", markdown=True, border=0)
    pdf.multi_cell(130, 5,
                   f"**\"Sectional Titles Schemes Management Act\"** means the Sectional Titles Schemes Management Act No. 8 of 2011 (or any statutory modification or re-enactment thereof) and includes the regulations made thereunder from time to time.",
                   align="J", markdown=True, new_x=XPos.LMARGIN, new_y=YPos.NEXT,
                   border=0)

    pdf.cell(20, 7, f"", align="C", markdown=True, border=0)
    pdf.cell(20, 7, f"", align="C", markdown=True, border=0)
    pdf.cell(20, 7, f"2.1.25", align="C", markdown=True, border=0)
    pdf.multi_cell(130, 5,
                   f"**\"Seller\"** means the party described in the Information Schedule. \n**\"Seller's attorneys\"** means Ilismi du Toit of LAÄS & SCHOLTZ INC, Queen Street Chambers, 33 Queen Street, Durbanville, 7550, Tel (021) 975 0802 \nEmail: ilismi@lslaw.co.za, LAÄS & SCHOLTZ INC Attorneys Trust Bank Account details;  \nBank:  Standard Bank; Account Number: 272255505;  Branch Code: 051001; (Ref: Unit Number / {data['development'].upper()})",
                   align="J", markdown=True, new_x=XPos.LMARGIN, new_y=YPos.NEXT,
                   border=0)

    pdf.cell(20, 7, f"", align="C", markdown=True, border=0)
    pdf.cell(20, 7, f"", align="C", markdown=True, border=0)
    pdf.cell(20, 7, f"2.1.26", align="C", markdown=True, border=0)
    pdf.multi_cell(130, 5,
                   f"**\"Signature Date\"** means the date upon which this Agreement is signed by the party who signs same last in time.",
                   align="J", markdown=True, new_x=XPos.LMARGIN, new_y=YPos.NEXT,
                   border=0)

    pdf.cell(20, 7, f"", align="C", markdown=True, border=0)
    pdf.cell(20, 7, f"", align="C", markdown=True, border=0)
    pdf.cell(20, 7, f"2.1.27", align="C", markdown=True, border=0)
    pdf.multi_cell(130, 5,
                   f"**\"Snags\"** means aesthetic and detail finishing items not affecting the Beneficial Occupation of the Property.",
                   align="J", markdown=True, new_x=XPos.LMARGIN, new_y=YPos.NEXT,
                   border=0)

    pdf.cell(20, 7, f"", align="C", markdown=True, border=0)
    pdf.cell(20, 7, f"", align="C", markdown=True, border=0)
    pdf.cell(20, 7, f"2.1.28", align="C", markdown=True, border=0)
    pdf.multi_cell(130, 5,
                   f"**\"Defects List\"** means a list furnished by the Seller to the Purchaser on Occupation Date, which list is to be completed by the Purchaser within 10 days after the Occupation Date, where the Purchaser may identify construction items inside the Property that are to be attended to by the Seller.",
                   align="J", markdown=True, new_x=XPos.LMARGIN, new_y=YPos.NEXT,
                   border=0)

    pdf.cell(20, 7, f"", align="C", markdown=True, border=0)
    pdf.cell(20, 7, f"", align="C", markdown=True, border=0)
    pdf.cell(20, 7, f"2.1.29", align="C", markdown=True, border=0)
    pdf.multi_cell(130, 5,
                   f"**\"Transfer Date\"** means the date of registration of transfer of the Unit in the name of the Purchaser in the deeds office.",
                   align="J", markdown=True, new_x=XPos.LMARGIN, new_y=YPos.NEXT,
                   border=0)

    pdf.cell(20, 7, f"", align="C", markdown=True, border=0)
    pdf.cell(20, 7, f"", align="C", markdown=True, border=0)
    pdf.cell(20, 7, f"2.1.30", align="C", markdown=True, border=0)
    pdf.multi_cell(130, 5,
                   f"**\"Unit\"** means the residential Sectional Title Unit to be constructed by the Seller on the Land for and on behalf of the Purchaser as envisaged herein.",
                   align="J", markdown=True, new_x=XPos.LMARGIN, new_y=YPos.NEXT,
                   border=0)

    pdf.cell(20, 7, f"", align="C", markdown=True, border=0)
    pdf.cell(20, 7, f"", align="C", markdown=True, border=0)
    pdf.cell(20, 7, f"2.1.31", align="C", markdown=True, border=0)
    pdf.multi_cell(130, 5,
                   f"**\"VAT\"** means value-added tax at the applicable rate in terms of the Value Added Tax Act No 89 of 1991 or any statutory re-enactment or amendment thereof.",
                   align="J", markdown=True, new_x=XPos.LMARGIN, new_y=YPos.NEXT,
                   border=0)

    pdf.cell(20, 7, f"", align="C", markdown=True, border=0)
    pdf.cell(20, 7, f"", align="C", markdown=True, border=0)
    pdf.cell(20, 7, f"2.1.32", align="C", markdown=True, border=0)
    pdf.multi_cell(130, 5,
                   f"**\"Works\"** means all the activities which are required to be undertaken to erect a residential Unit on the Land for purposes of handover and transfer to the Purchaser.",
                   align="J", markdown=True, new_x=XPos.LMARGIN, new_y=YPos.NEXT,
                   border=0)

    pdf.cell(20, 7, f"", align="C", markdown=True, border=0)
    pdf.cell(20, 7, f"2.2", align="C", markdown=True, border=0)
    pdf.multi_cell(150, 5, f"The headnotes to the paragraphs in this Agreement are inserted for reference purposes only and shall not affect the interpretation of any of the provisions to which they relate.",
                   align="J", markdown=True, new_x=XPos.LMARGIN, new_y=YPos.NEXT, border=0)

    pdf.cell(20, 7, f"", align="C", markdown=True, border=0)
    pdf.cell(20, 7, f"2.3", align="C", markdown=True, border=0)
    pdf.multi_cell(150, 5,
                   f"Words importing the singular shall include the plural, and vice versa, and words importing the masculine gender shall include the feminine and neuter genders, and vice versa, and words importing persons shall include partnerships, trusts and bodies corporate, and *vice versa*.",
                   align="J", markdown=True, new_x=XPos.LMARGIN, new_y=YPos.NEXT, border=0)


    pdf.add_page()


    pdf.cell(20, 7, f"", align="C", markdown=True, border=0)
    pdf.cell(20, 7, f"2.4", align="C", markdown=True, border=0)
    pdf.multi_cell(150, 5,
                   f"If any provision in the Information Schedule, clause 1 and/or this clause 2 is a substantive provision conferring rights or imposing obligations on any party, then notwithstanding that such provision is contained in the Information Schedule, Clause 1 and/or this Clause 2  (as the case may be) effect shall be given thereto as if such provision was a substantive provision in the body of this Agreement.",
                   align="J", markdown=True, new_x=XPos.LMARGIN, new_y=YPos.NEXT, border=0)

    pdf.cell(20, 7, f"3.", align="C", markdown=True, border=0)
    pdf.cell(50, 7, f"**SALE OF THE PROPERTY**", align="L", markdown=True, border=0, new_x=XPos.LMARGIN,
             new_y=YPos.NEXT)

    pdf.cell(20, 7, f"", align="C", markdown=True, border=0)
    pdf.multi_cell(170, 7, f"The Seller hereby sells, and the Purchaser hereby purchases the Property, subject to and upon the terms and conditions contained in this Agreement.", align="L", markdown=True, border=0, new_x=XPos.LMARGIN,
             new_y=YPos.NEXT)

    pdf.cell(20, 7, f"4.", align="C", markdown=True, border=0)
    pdf.cell(50, 7, f"**PURCHASE PRICE AND METHOD OF PAYMENT**", align="L", markdown=True, border=0, new_x=XPos.LMARGIN,
             new_y=YPos.NEXT)

    pdf.cell(20, 7, f"", align="C", markdown=True, border=0)
    pdf.cell(20, 7, f"4.1", align="C", markdown=True, border=0)
    pdf.multi_cell(150, 5, f"The Total purchase price of the Property shall be the amount stated in clause E1 of the Information Schedule regardless of the final extent of the Unit as reflected on the Unit Plan attached marked \"Annexure B\"",
                   align="J", markdown=True, new_x=XPos.LMARGIN, new_y=YPos.NEXT, border=0)

    pdf.cell(20, 7, f"", align="C", markdown=True, border=0)
    pdf.cell(20, 7, f"", align="C", markdown=True, border=0)
    pdf.cell(20, 7, f"4.1.1", align="C", markdown=True, border=0)
    pdf.multi_cell(130, 5,
                   f"The Purchaser shall pay the Seller's attorneys the deposit for the Property as stated in clause E2 of the Information Schedule within 3(three) days of signature of this Agreement by the Purchaser, which deposit shall be held in trust by the Seller's attorneys and invested in an interest-bearing account in accordance with the provisions of Section 26 of the Alienation of Land Act No 68 of 1981 (as amended) with interest to accrue to the Purchaser.  The provisions of this clause 4.2 shall constitute authority to the Seller's attorneys, in terms of Section 86(4) of the Legal Practice Act, 2014(Act No. 28 of 2014), to invest the deposit for the benefit of Purchaser pending registration of transfer.",
                   align="J", markdown=True, new_x=XPos.LMARGIN, new_y=YPos.NEXT,
                   border=0)

    pdf.cell(20, 7, f"", align="C", markdown=True, border=0)
    pdf.cell(20, 7, f"", align="C", markdown=True, border=0)
    pdf.cell(20, 7, f"4.1.2", align="C", markdown=True, border=0)
    pdf.multi_cell(130, 5,
                   f"Authority is hereby granted to the attorney by the Purchaser to withdraw the deposit as provided for in clause 20.4 of the Agreement.",
                   align="J", markdown=True, new_x=XPos.LMARGIN, new_y=YPos.NEXT,
                   border=0)

    pdf.cell(20, 7, f"", align="C", markdown=True, border=0)
    pdf.cell(20, 7, f"", align="C", markdown=True, border=0)
    pdf.cell(20, 7, f"4.1.3", align="C", markdown=True, border=0)
    pdf.multi_cell(130, 5,
                   f"The Purchaser is aware that upon such withdrawal, no interest will be earned on the portion being withdrawn.",
                   align="J", markdown=True, new_x=XPos.LMARGIN, new_y=YPos.NEXT,
                   border=0)

    pdf.cell(20, 7, f"", align="C", markdown=True, border=0)
    pdf.cell(20, 7, f"4.2", align="C", markdown=True, border=0)
    pdf.multi_cell(150, 5,
                   f"**The Seller will not be bound to the Purchaser in terms of this Agreement until such time as the deposit referred to in clause E2 has been paid to the Sellers attorneys trust account** referred to in clause 4.1 above. The Seller shall be entitled to accept further offers acceptable to the Seller, until such time as proof of payment of the deposit is furnished to the Seller or the Seller's Attorneys, by the Purchaser, as provided for in this clause 4.2 In the event of the Seller accepting an offer to purchase the Property on terms and conditions acceptable to the Seller prior to receipt of such written notification, this Agreement shall be deemed ipso facto null and void.",
                   align="J", markdown=True, new_x=XPos.LMARGIN, new_y=YPos.NEXT, border=0)

    pdf.cell(20, 7, f"", align="C", markdown=True, border=0)
    pdf.cell(20, 7, f"4.3", align="C", markdown=True, border=0)
    pdf.multi_cell(150, 5,
                   f"Within **21 (twenty one)** days after signature of this Agreement, the Purchaser shall furnish the Seller or the Seller's Attorneys, with an irrevocable guarantee issued by a registered commercial bank for the due payment of the balance of the purchase price of the Property as referred to in clause E3 of the Information Schedule, or in the event of the Purchaser requiring a mortgage bond for purposes of purchasing the Property in the amount recorded in clause E3 of the Information Schedule, within **21 (twenty one)** days of securing a mortgage bond as provided for in clause 19.1 hereunder. Should the Purchaser fail to comply with this clause 4.4, the contract will be deemed null and void.",
                   align="J", markdown=True, new_x=XPos.LMARGIN, new_y=YPos.NEXT, border=0)

    pdf.cell(20, 7, f"", align="C", markdown=True, border=0)
    pdf.cell(20, 7, f"4.4", align="C", markdown=True, border=0)
    pdf.multi_cell(150, 5,
                   f"Or alternatively to the delivery of the guarantee referred to in clause 4.4 above, the Purchaser shall within the same time periods as provided for in the aforesaid clause, pay into the trust account of the Seller's attorneys, the balance of the purchase price of the Property as referred to in clause E3 of the Information Schedule, to be held by such attorneys in an interest bearing trust account, interest to accrue for the benefit of the Purchaser until the date upon which payment of the relevant amount falls due to the Seller.  The Purchaser hereby irrevocably authorises the attorneys to release from the funds so received, the payments due to the Seller in terms of the provisions of this Agreement.",
                   align="J", markdown=True, new_x=XPos.LMARGIN, new_y=YPos.NEXT, border=0)










    pdf.output(f"sales_client_onboarding_docs/{data['opportunity_code']}-OTP.pdf")

    return f"sales_client_onboarding_docs/{data['opportunity_code']}-OTP.pdf"
