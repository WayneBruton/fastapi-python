def create_signatories(data, number_of_purchasers):
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
    return purchaser_signing_details
