import copy
import os

from fpdf import FPDF, XPos, YPos
import PyPDF2


class PDF(FPDF):
    def footer(self):
        self.set_y(-15)
        self.set_font('helvetica', 'I', 6)
        self.cell(0, 10, f"Page {self.page_no()}/{{nb}}", align="C")


def print_onboarding_pdf(data):
    # print("DATA", data)
    pdf = PDF('P', 'mm', 'A4')
    pdf.set_auto_page_break(auto=True, margin=15)
    pdf.add_page()
    # pdf.set_font('helvetica', '', 10)
    pdf.set_fill_color(211, 211, 211)
    pdf.set_font('helvetica', '', 12)

    # unit_name = 'EA107'

    pdf.cell(0, 5, f"**{data['opportunity_code']} - INFORMATION SHEET**", new_x=XPos.LMARGIN, new_y=YPos.NEXT,
             markdown=True,
             align="C")

    pdf.cell(0, 5, f"**{data['development']}**", new_x=XPos.LMARGIN, new_y=YPos.NEXT, markdown=True,
             align="C")
    pdf.set_font('helvetica', '', 10)
    pdf.cell(0, 5, "", new_x=XPos.LMARGIN, new_y=YPos.NEXT)
    pdf.cell(0, 5, "", new_x=XPos.LMARGIN, new_y=YPos.NEXT)
    pdf.cell(0, 5, "", new_x=XPos.LMARGIN, new_y=YPos.NEXT)

    pdf.cell(90, 5, f"**Client Type: **", markdown=True,
             align="L", border=True)
    pdf.cell(5, 5, "", border=False)
    pdf.cell(95, 5, f"**{data['opportunity_client_type']}**", new_x=XPos.LMARGIN, new_y=YPos.NEXT, markdown=True,
             align="L", border=True)

    if data['opportunity_companyname'] is not None and data['opportunity_companyname'] != "":
        pdf.cell(0, 3, "", new_x=XPos.LMARGIN, new_y=YPos.NEXT)
        pdf.cell(90, 5, f"**Company / Trust Name: **", markdown=True,
                 align="L", border=True)
        pdf.cell(5, 5, "", border=False)
        pdf.cell(95, 5, f"**{data['opportunity_companyname']}**", new_x=XPos.LMARGIN, new_y=YPos.NEXT, markdown=True,
                 align="L", border=True)

    pdf.cell(0, 3, "", new_x=XPos.LMARGIN, new_y=YPos.NEXT)
    pdf.cell(90, 5, f"**Agent: **", markdown=True,
             align="L", border=True)
    pdf.cell(5, 5, "", border=False)
    pdf.cell(95, 5, f"**{data['opportunity_sale_agreement']}**", new_x=XPos.LMARGIN, new_y=YPos.NEXT, markdown=True,
             align="L", border=True)

    pdf.cell(0, 3, "", new_x=XPos.LMARGIN, new_y=YPos.NEXT)
    pdf.cell(90, 5, f"**No of Purchasers: **", markdown=True,
             align="L", border=True)
    pdf.cell(5, 5, "", border=False)
    pdf.cell(95, 5, f"**{data['opportunity_client_no']}**", new_x=XPos.LMARGIN, new_y=YPos.NEXT, markdown=True,
             align="L", border=True)

    number_of_purchasers = int(data['opportunity_client_no'].split(" ")[0])
    # print("number_of_purchasers", number_of_purchasers)

    basic_info_A = {
        'info1': "",
        'opportunity_martial_status': 'opportunity_martial_status',
        'opportunity_firstname': 'opportunity_firstname',
        'opportunity_lastname': 'opportunity_lastname',
        'opportunity_id': 'opportunity_id',
        'opportunity_email': 'opportunity_email',
        'opportunity_mobile': 'opportunity_mobile',
        'opportunity_landline': 'opportunity_landline',
        'opportunity_postal_address': 'opportunity_postal_address',
        'opportunity_residental_address': 'opportunity_residental_address',
    }
    # make a deepcopy of basic_info
    basic_info = copy.deepcopy(basic_info_A)

    purchasers = []
    if number_of_purchasers >= 1:
        basic_info = copy.deepcopy(basic_info_A)
        basic_info['info1'] = "Details of 1st Person"
        purchasers.append(basic_info)
    if number_of_purchasers >= 2:
        basic_info = copy.deepcopy(basic_info_A)
        basic_info['info1'] = "Details of 2nd Person"
        # loop through basic_info and change the values
        for key in basic_info:
            if not basic_info[key] == "info1":
                basic_info[key] = basic_info[key] + "_sec"
        purchasers.append(basic_info)
    if number_of_purchasers >= 3:
        basic_info = copy.deepcopy(basic_info_A)
        basic_info['info1'] = "Details of 3rd Person"
        # loop through basic_info and change the values
        for key in basic_info:
            basic_info = copy.deepcopy(basic_info_A)
            if not basic_info[key] == "info1":
                basic_info[key] = basic_info[key] + "_3rd"
        purchasers.append(basic_info)
    if number_of_purchasers >= 4:
        basic_info = copy.deepcopy(basic_info_A)
        basic_info['info1'] = "Details of 4th Person"
        # loop through basic_info and change the values
        for key in basic_info:
            if not basic_info[key] == "info1":
                basic_info[key] = basic_info[key] + "_4th"
        purchasers.append(basic_info)
    if number_of_purchasers >= 5:
        basic_info = copy.deepcopy(basic_info_A)
        basic_info['info1'] = "Details of 5th Person"
        # loop through basic_info and change the values
        for key in basic_info:
            if not basic_info[key] == "info1":
                basic_info[key] = basic_info[key] + "_5th"
        purchasers.append(basic_info)
    if number_of_purchasers >= 6:
        basic_info = copy.deepcopy(basic_info_A)
        basic_info['info1'] = "Details of 6th Person"
        # loop through basic_info and change the values
        for key in basic_info:
            if not basic_info[key] == "info1":
                basic_info[key] = basic_info[key] + "_6th"
        purchasers.append(basic_info)
    if number_of_purchasers >= 7:
        basic_info = copy.deepcopy(basic_info_A)
        basic_info['info1'] = "Details of 7th Person"
        # loop through basic_info and change the values
        for key in basic_info:
            if not basic_info[key] == "info1":
                basic_info[key] = basic_info[key] + "_7th"
        purchasers.append(basic_info)
    if number_of_purchasers >= 8:
        basic_info = copy.deepcopy(basic_info_A)
        basic_info['info1'] = "Details of 8th Person"
        # loop through basic_info and change the values
        for key in basic_info:
            if not basic_info[key] == "info1":
                basic_info[key] = basic_info[key] + "_8th"
        purchasers.append(basic_info)
    if number_of_purchasers >= 9:
        basic_info = copy.deepcopy(basic_info_A)
        basic_info['info1'] = "Details of 9th Person"
        # loop through basic_info and change the values
        for key in basic_info:
            if not basic_info[key] == "info1":
                basic_info[key] = basic_info[key] + "_9th"
        purchasers.append(basic_info)
    if number_of_purchasers >= 10:
        basic_info = copy.deepcopy(basic_info_A)
        basic_info['info1'] = "Details of 10th Person"
        # loop through basic_info and change the values
        for key in basic_info:
            if not basic_info[key] == "info1":
                basic_info[key] = basic_info[key] + "_10th"
        purchasers.append(basic_info)

    # print("purchasers", purchasers[3])
    # print("purchasers", len(purchasers))

    for purchaser in purchasers:
        pdf.cell(0, 3, "", new_x=XPos.LMARGIN, new_y=YPos.NEXT, border="B")
        pdf.cell(0, 3, "", new_x=XPos.LMARGIN, new_y=YPos.NEXT)
        pdf.cell(0, 5, f"**{purchaser['info1']}**", markdown=True,
                 align="C", border=False, new_x=XPos.LMARGIN, new_y=YPos.NEXT)
        pdf.cell(0, 3, "", new_x=XPos.LMARGIN, new_y=YPos.NEXT)
        pdf.cell(90, 5, f"**Marital Status: **", markdown=True,
                 align="L", border=True)
        pdf.cell(5, 5, "", border=False)
        pdf.cell(95, 5, f"**{data[purchaser['opportunity_martial_status']]}**", new_x=XPos.LMARGIN, new_y=YPos.NEXT,
                 markdown=True,
                 align="L", border=True)

        pdf.cell(0, 3, "", new_x=XPos.LMARGIN, new_y=YPos.NEXT)
        pdf.cell(90, 5, f"**First Name: **", markdown=True,
                 align="L", border=True)
        pdf.cell(5, 5, "", border=False)
        pdf.cell(95, 5, f"**{data[purchaser['opportunity_firstname']]}**", new_x=XPos.LMARGIN, new_y=YPos.NEXT,
                 markdown=True,
                 align="L", border=True)

        pdf.cell(0, 3, "", new_x=XPos.LMARGIN, new_y=YPos.NEXT)
        pdf.cell(90, 5, f"**Surname: **", markdown=True,
                 align="L", border=True)
        pdf.cell(5, 5, "", border=False)
        pdf.cell(95, 5, f"**{data[purchaser['opportunity_lastname']]}**", new_x=XPos.LMARGIN, new_y=YPos.NEXT,
                 markdown=True,
                 align="L", border=True)

        pdf.cell(0, 3, "", new_x=XPos.LMARGIN, new_y=YPos.NEXT)
        pdf.cell(90, 5, f"**ID Number: **", markdown=True,
                 align="L", border=True)
        pdf.cell(5, 5, "", border=False)
        pdf.cell(95, 5, f"**{data[purchaser['opportunity_id']]}**", new_x=XPos.LMARGIN, new_y=YPos.NEXT, markdown=True,
                 align="L", border=True)

        pdf.cell(0, 3, "", new_x=XPos.LMARGIN, new_y=YPos.NEXT)
        pdf.cell(90, 5, f"**Email: **", markdown=True,
                 align="L", border=True)
        pdf.cell(5, 5, "", border=False)
        pdf.cell(95, 5, f"**{data[purchaser['opportunity_email']]}**", new_x=XPos.LMARGIN, new_y=YPos.NEXT,
                 markdown=True,
                 align="L", border=True)

        pdf.cell(0, 3, "", new_x=XPos.LMARGIN, new_y=YPos.NEXT)
        pdf.cell(90, 5, f"**Mobile: **", markdown=True,
                 align="L", border=True)
        pdf.cell(5, 5, "", border=False)
        pdf.cell(95, 5, f"**{data[purchaser['opportunity_mobile']]}**", new_x=XPos.LMARGIN, new_y=YPos.NEXT,
                 markdown=True,
                 align="L", border=True)

        pdf.cell(0, 3, "", new_x=XPos.LMARGIN, new_y=YPos.NEXT)
        pdf.cell(90, 5, f"**Landline: **", markdown=True,
                 align="L", border=True)
        pdf.cell(5, 5, "", border=False)
        pdf.cell(95, 5, f"**{data[purchaser['opportunity_landline']]}**", new_x=XPos.LMARGIN, new_y=YPos.NEXT,
                 markdown=True,
                 align="L", border=True)

        pdf.cell(0, 3, "", new_x=XPos.LMARGIN, new_y=YPos.NEXT)
        pdf.cell(90, 5, f"**Postal Address: **", markdown=True,
                 align="L", border=True)
        pdf.cell(5, 5, "", border=False)
        pdf.cell(95, 5, f"**{data[purchaser['opportunity_postal_address']]}**", new_x=XPos.LMARGIN, new_y=YPos.NEXT,
                 markdown=True,
                 align="L", border=True)

        pdf.cell(0, 3, "", new_x=XPos.LMARGIN, new_y=YPos.NEXT)
        pdf.cell(90, 5, f"**Residential Address: **", markdown=True,
                 align="L", border=True)
        pdf.cell(5, 5, "", border=False)
        pdf.cell(95, 5, f"**{data[purchaser['opportunity_residental_address']]}**", new_x=XPos.LMARGIN, new_y=YPos.NEXT,
                 markdown=True,
                 align="L", border=True)

    pdf.cell(0, 3, "", new_x=XPos.LMARGIN, new_y=YPos.NEXT)
    pdf.cell(0, 3, "", new_x=XPos.LMARGIN, new_y=YPos.NEXT, border="B")

    pdf.add_page()

    pdf.set_font('helvetica', '', 12)

    pdf.cell(0, 5, f"**{data['opportunity_code']} - PURCHASE DETAILS**", new_x=XPos.LMARGIN, new_y=YPos.NEXT,
             markdown=True,
             align="C")

    pdf.set_font('helvetica', '', 10)
    pdf.cell(0, 5, "", new_x=XPos.LMARGIN, new_y=YPos.NEXT)
    pdf.cell(0, 5, "", new_x=XPos.LMARGIN, new_y=YPos.NEXT)
    pdf.cell(0, 5, "", new_x=XPos.LMARGIN, new_y=YPos.NEXT)

    if 'opportunity_pay_type' not in data:
        data['opportunity_pay_type'] = None

    pdf.cell(90, 5, f"**Purchase Type: **", markdown=True,
             align="L", border=True)
    pdf.cell(5, 5, "", border=False)
    pdf.cell(95, 5, f"**{data['opportunity_pay_type']}**", new_x=XPos.LMARGIN, new_y=YPos.NEXT, markdown=True,
             align="L", border=True)

    if 'opportunity_sales_date' not in data:
        data['opportunity_sales_date'] = None

    pdf.cell(0, 3, "", new_x=XPos.LMARGIN, new_y=YPos.NEXT)
    pdf.cell(90, 5, f"**Sales Date: **", markdown=True,
             align="L", border=True)
    pdf.cell(5, 5, "", border=False)
    pdf.cell(95, 5, f"**{data['opportunity_sales_date']}**", new_x=XPos.LMARGIN, new_y=YPos.NEXT, markdown=True,
             align="L", border=True)

    if 'opportunity_base_price' not in data:
        data['opportunity_base_price'] = None

    pdf.cell(0, 3, "", new_x=XPos.LMARGIN, new_y=YPos.NEXT)
    pdf.cell(90, 5, f"**Base Price: **", markdown=True,
             align="L", border=True)
    pdf.cell(5, 5, "", border=False)
    pdf.cell(95, 5, f"**{data['opportunity_base_price']}**", new_x=XPos.LMARGIN, new_y=YPos.NEXT, markdown=True,
             align="L", border=True)

    if 'opportunity_discount' not in data:
        data['opportunity_discount'] = None

    pdf.cell(0, 3, "", new_x=XPos.LMARGIN, new_y=YPos.NEXT)
    pdf.cell(90, 5, f"**Discount: **", markdown=True,
             align="L", border=True)
    pdf.cell(5, 5, "", border=False)
    pdf.cell(95, 5, f"**{data['opportunity_discount']}**", new_x=XPos.LMARGIN, new_y=YPos.NEXT, markdown=True,
             align="L", border=True)

    if 'opportunity_deposite_date' not in data:
        data['opportunity_deposite_date'] = None

    pdf.cell(0, 3, "", new_x=XPos.LMARGIN, new_y=YPos.NEXT)
    pdf.cell(90, 5, f"**Deposit Date: **", markdown=True,
             align="L", border=True)
    pdf.cell(5, 5, "", border=False)
    pdf.cell(95, 5, f"**{data['opportunity_deposite_date']}**", new_x=XPos.LMARGIN, new_y=YPos.NEXT, markdown=True,
             align="L", border=True)

    if 'opportunity_deposite' not in data:
        data['opportunity_deposite'] = None

    pdf.cell(0, 3, "", new_x=XPos.LMARGIN, new_y=YPos.NEXT)
    pdf.cell(90, 5, f"**Deposit: **", markdown=True,
             align="L", border=True)
    pdf.cell(5, 5, "", border=False)
    pdf.cell(95, 5, f"**{data['opportunity_deposite']}**", new_x=XPos.LMARGIN, new_y=YPos.NEXT, markdown=True,
             align="L", border=True)

    if 'opportunity_originalBayNo' not in data:
        data['opportunity_originalBayNo'] = None

    pdf.cell(0, 3, "", new_x=XPos.LMARGIN, new_y=YPos.NEXT)
    pdf.cell(90, 5, f"**Allocated Bay No: **", markdown=True,
             align="L", border=True)
    pdf.cell(5, 5, "", border=False)
    pdf.cell(95, 5, f"**{data['opportunity_originalBayNo']}**", new_x=XPos.LMARGIN, new_y=YPos.NEXT, markdown=True,
             align="L", border=True)

    if 'opportunity_additional_bay_covered' not in data:
        data['opportunity_additional_bay_covered'] = None

    pdf.cell(0, 3, "", new_x=XPos.LMARGIN, new_y=YPos.NEXT)
    pdf.cell(90, 5, f"**Additional Bay Type: **", markdown=True,
             align="L", border=True)
    pdf.cell(5, 5, "", border=False)
    pdf.cell(95, 5, f"**{data['opportunity_additional_bay_covered']}**", new_x=XPos.LMARGIN, new_y=YPos.NEXT,
             markdown=True,
             align="L", border=True)

    if 'opportunity_additional_bay_free' not in data:
        data['opportunity_additional_bay_free'] = None

    pdf.cell(0, 3, "", new_x=XPos.LMARGIN, new_y=YPos.NEXT)
    pdf.cell(90, 5, f"**Free Parking: **", markdown=True,
             align="L", border=True)
    pdf.cell(5, 5, "", border=False)
    pdf.cell(95, 5, f"**{data['opportunity_additional_bay_free']}**", new_x=XPos.LMARGIN, new_y=YPos.NEXT,
             markdown=True,
             align="L", border=True)

    if 'opportunity_additional_bay' not in data:
        data['opportunity_additional_bay'] = None

    pdf.cell(0, 3, "", new_x=XPos.LMARGIN, new_y=YPos.NEXT)
    pdf.cell(90, 5, f"**Additional Bays: **", markdown=True,
             align="L", border=True)
    pdf.cell(5, 5, "", border=False)
    pdf.cell(95, 5, f"**{data['opportunity_additional_bay']}**", new_x=XPos.LMARGIN, new_y=YPos.NEXT,
             markdown=True,
             align="L", border=True)

    if 'opportunity_parking_base' not in data:
        data['opportunity_parking_base'] = None

    pdf.cell(0, 3, "", new_x=XPos.LMARGIN, new_y=YPos.NEXT)
    pdf.cell(90, 5, f"**Parking Bay Number: **", markdown=True,
             align="L", border=True)
    pdf.cell(5, 5, "", border=False)
    pdf.cell(95, 5, f"**{data['opportunity_parking_base']}**", new_x=XPos.LMARGIN, new_y=YPos.NEXT,
             markdown=True,
             align="L", border=True)

    if 'opportunity_parking_base2' not in data:
        data['opportunity_parking_base2'] = None

    pdf.cell(0, 3, "", new_x=XPos.LMARGIN, new_y=YPos.NEXT)
    pdf.cell(90, 5, f"**Parking Bay Number: **", markdown=True,
             align="L", border=True)
    pdf.cell(5, 5, "", border=False)
    pdf.cell(95, 5, f"**{data['opportunity_parking_base2']}**", new_x=XPos.LMARGIN, new_y=YPos.NEXT,
             markdown=True,
             align="L", border=True)

    if 'opportunity_parking_cost' not in data:
        data['opportunity_parking_cost'] = None

    pdf.cell(0, 3, "", new_x=XPos.LMARGIN, new_y=YPos.NEXT)
    pdf.cell(90, 5, f"**Parking Bay Cost: **", markdown=True,
             align="L", border=True)
    pdf.cell(5, 5, "", border=False)
    pdf.cell(95, 5, f"**{data['opportunity_parking_cost']}**", new_x=XPos.LMARGIN, new_y=YPos.NEXT,
             markdown=True,
             align="L", border=True)

    if 'opportunity_stove_option' not in data:
        data['opportunity_stove_option'] = None

    pdf.cell(0, 3, "", new_x=XPos.LMARGIN, new_y=YPos.NEXT)
    pdf.cell(90, 5, f"**Stove Option: **", markdown=True,
             align="L", border=True)
    pdf.cell(5, 5, "", border=False)
    pdf.cell(95, 5, f"**{data['opportunity_stove_option']}**", new_x=XPos.LMARGIN, new_y=YPos.NEXT,
             markdown=True,
             align="L", border=True)

    if 'opportunity_stove_cost' not in data:
        data['opportunity_stove_cost'] = None

    pdf.cell(0, 3, "", new_x=XPos.LMARGIN, new_y=YPos.NEXT)
    pdf.cell(90, 5, f"**Stove Cost: **", markdown=True,
             align="L", border=True)
    pdf.cell(5, 5, "", border=False)
    pdf.cell(95, 5, f"**{data['opportunity_stove_cost']}**", new_x=XPos.LMARGIN, new_y=YPos.NEXT,
             markdown=True,
             align="L", border=True)

    if 'opportunity_additional_cost' not in data:
        data['opportunity_additional_cost'] = None

    pdf.cell(0, 3, "", new_x=XPos.LMARGIN, new_y=YPos.NEXT)
    pdf.cell(90, 5, f"**Additional Extra Cost: **", markdown=True,
             align="L", border=True)
    pdf.cell(5, 5, "", border=False)
    pdf.cell(95, 5, f"**{data['opportunity_additional_cost']}**", new_x=XPos.LMARGIN, new_y=YPos.NEXT,
             markdown=True,
             align="L", border=True)

    if 'opportunity_bond_amount' not in data:
        data['opportunity_bond_amount'] = None

    pdf.cell(0, 3, "", new_x=XPos.LMARGIN, new_y=YPos.NEXT)
    pdf.cell(90, 5, f"**Bond Amount Required: **", markdown=True,
             align="L", border=True)
    pdf.cell(5, 5, "", border=False)
    pdf.cell(95, 5, f"**{data['opportunity_bond_amount']}**", new_x=XPos.LMARGIN, new_y=YPos.NEXT,
             markdown=True,
             align="L", border=True)

    if 'opportunity_contract_price' not in data:
        data['opportunity_contract_price'] = None

    pdf.cell(0, 3, "", new_x=XPos.LMARGIN, new_y=YPos.NEXT)
    pdf.cell(90, 5, f"**Contract Price: **", markdown=True,
             align="L", border=True)
    pdf.cell(5, 5, "", border=False)
    pdf.cell(95, 5, f"**{data['opportunity_contract_price']}**", new_x=XPos.LMARGIN, new_y=YPos.NEXT,
             markdown=True,
             align="L", border=True)

    if 'opportunity_gardenNumber' not in data:
        data['opportunity_gardenNumber'] = None

    pdf.cell(0, 3, "", new_x=XPos.LMARGIN, new_y=YPos.NEXT)
    pdf.cell(90, 5, f"**Garden Number: **", markdown=True,
             align="L", border=True)
    pdf.cell(5, 5, "", border=False)
    pdf.cell(95, 5, f"**{data['opportunity_gardenNumber']}**", new_x=XPos.LMARGIN, new_y=YPos.NEXT,
             markdown=True,
             align="L", border=True)

    if 'opportunity_gardenSize' not in data:
        data['opportunity_gardenSize'] = None

    pdf.cell(0, 3, "", new_x=XPos.LMARGIN, new_y=YPos.NEXT)
    pdf.cell(90, 5, f"**Garden Size: **", markdown=True,
             align="L", border=True)
    pdf.cell(5, 5, "", border=False)
    pdf.cell(95, 5, f"**{data['opportunity_gardenSize']}**", new_x=XPos.LMARGIN, new_y=YPos.NEXT,
             markdown=True,
             align="L", border=True)

    if 'opportunity_specials' not in data:
        data['opportunity_specials'] = []
    print("TEST")

    if data['opportunity_specials'] != None and len(data['opportunity_specials']) > 0:

        for index, special in enumerate(data['opportunity_specials'], 1):
            pdf.cell(0, 3, "", new_x=XPos.LMARGIN, new_y=YPos.NEXT)
            pdf.cell(90, 5, f"**Included Specials No {index}: **", markdown=True,
                     align="L", border=True)
            pdf.cell(5, 5, "", border=False)
            pdf.cell(95, 5, f"**{special}**", new_x=XPos.LMARGIN, new_y=YPos.NEXT,
                     markdown=True,
                     align="L", border=True)

    if 'opportunity_notes' not in data:
        data['opportunity_notes'] = None



    pdf.cell(0, 3, "", new_x=XPos.LMARGIN, new_y=YPos.NEXT)
    pdf.cell(90, 5, f"**Notes: **", markdown=True,
             align="L", border=True)
    pdf.cell(5, 5, "", border=False)
    pdf.cell(95, 5, f"**{data['opportunity_notes']}**", new_x=XPos.LMARGIN, new_y=YPos.NEXT,
             markdown=True,
             align="L", border=True)

    if 'opportunity_extras_not_listed' not in data:
        data['opportunity_extras_not_listed'] = []

    # delete the first record in data['opportunity_extras_not_listed'] list
    data['opportunity_extras_not_listed'].pop(0)

    if len(data['opportunity_extras_not_listed']) > 0:
        for extra in data['opportunity_extras_not_listed']:
            pdf.cell(0, 3, "", new_x=XPos.LMARGIN, new_y=YPos.NEXT)
            pdf.cell(90, 5, f"**Amount - Extras -  ({extra['description']}): **", markdown=True,
                     align="L", border=True)
            pdf.cell(5, 5, "", border=False)
            pdf.cell(95, 5, f"**{extra['value']}**", new_x=XPos.LMARGIN, new_y=YPos.NEXT,
                     markdown=True,
                     align="L", border=True)

    if 'opportunity_mood' not in data:
        data['opportunity_mood'] = None

    pdf.cell(0, 3, "", new_x=XPos.LMARGIN, new_y=YPos.NEXT)
    pdf.cell(90, 5, f"**Mood: **", markdown=True,
             align="L", border=True)
    pdf.cell(5, 5, "", border=False)
    pdf.cell(95, 5, f"**{data['opportunity_mood']}**", new_x=XPos.LMARGIN, new_y=YPos.NEXT,
             markdown=True,
             align="L", border=True)

    pdf.output(f"sales_client_onboarding_docs/{data['opportunity_code']}-onboarding.pdf")

    directory = 'sales_documents'
    filename = f"{data['opportunity_code']}-Annexure.pdf"

    file_path = os.path.join(directory, filename)

    if os.path.isfile(file_path):
        print(f'The file {filename} exists in the directory {directory}.')
        pdf1_file = open(f"sales_client_onboarding_docs/{data['opportunity_code']}-onboarding.pdf", 'rb')
        pdf1_reader = PyPDF2.PdfFileReader(pdf1_file)

        # Open the second PDF file in read-binary mode
        pdf2_file = open(f"sales_documents/{data['opportunity_code']}-Annexure.pdf", 'rb')
        pdf2_reader = PyPDF2.PdfFileReader(pdf2_file)

        # Create a new PDF writer object
        pdf_writer = PyPDF2.PdfFileWriter()

        # Append each page from the first PDF to the writer
        for page_num in range(pdf1_reader.numPages):
            page = pdf1_reader.getPage(page_num)
            pdf_writer.addPage(page)

        # Append each page from the second PDF to the writer
        for page_num in range(pdf2_reader.numPages):
            page = pdf2_reader.getPage(page_num)
            pdf_writer.addPage(page)

        # Save the combined PDF to a new file
        output_file = open(f"sales_client_onboarding_docs/{data['opportunity_code']}-onboarding.pdf", 'wb')
        pdf_writer.write(output_file)

        # Close the input and output files
        pdf1_file.close()
        pdf2_file.close()
        output_file.close()
    else:
        print(f'The file {filename} does not exist in the directory {directory}.')

    print(f"{data['opportunity_code']}-onboarding.pdf")

    return f"{data['opportunity_code']}-onboarding.pdf"
