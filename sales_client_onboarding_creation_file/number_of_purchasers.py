def create_purchaser_details(data, number_of_purchasers):
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
    return purchaser_details
