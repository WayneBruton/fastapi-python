from datetime import date, datetime

from fpdf import FPDF, XPos, YPos
import os


def create_pdf(statement_type, data):
    data1 = []
    data1 = data
    # print(data1)

    filename = f"{data1[0]['investor_acc_number']}_{data1[0]['opportunity_code']}_" \
               f"{data1[0]['investment_number']}_type{statement_type}.pdf"

    # if filename exists in portal_statements folder, delete it
    if os.path.exists(f"portal_statements/{filename}"):
        os.remove(f"portal_statements/{filename}")
        print(f"removed {filename}")

    today = date.today()
    # today's month
    month = today.strftime("%B")
    # today's year
    year = today.strftime("%Y")

    day = today.strftime("%d")

    today = f"{day} {month} {year}"
    for item in data:
        if 'days' not in item:
            item['days'] = ""

    if not 'investment_name' in data1[0]:
        data1[0]['investment_name'] = f"{data1[0]['investor_name']} {data1[0]['investor_surname']}"
    else:
        data1[0]['investment_name'] = f"{data1[0]['investment_name']}"

    front_page_data = [
        {
            "title": "**DEVELOPMENT**",
            "value": f"{data1[0]['Category']}",
        },
        {
            "title": "**BORROWER NAME**",
            "value": "Heron Projects",
        },
        {
            "title": "**LENDOR NAME**",
            # "value": f"{data1[0]['investor_name']} {data1[0]['investor_surname']} ({data1[0]['investor_acc_number']})",
            "value": f"{data1[0]['investment_name']}",
        },
        {
            "title": "**PROPERTY CODE**",
            "value": f"{data1[0]['opportunity_code']}",
        },
        {
            "title": "**STATEMENT DATE**",
            "value": f"{today}",
        },
        {
            "title": "**CURRENCY**",
            "value": "SA Rand",
        }
    ]

    pdf = FPDF()

    pdf.add_page()
    # FRONT PAGE
    # Add the image
    pdf.image('portal_statement_files/statement-logo.png', x=10, y=10, w=95)
    pdf.set_font('helvetica', size=12)
    pdf.cell(105, 10, '', border=0)
    pdf.cell(80, 10, "**LENDER STATEMENT**", border=0, new_x=XPos.LMARGIN, new_y=YPos.NEXT, align="C", markdown=True)

    pdf.cell(120, 5, '', border=0)
    pdf.cell(50, 5, "", border='B', new_x=XPos.LMARGIN, new_y=YPos.NEXT, align="C", markdown=True)
    pdf.cell(105, 5, '', border=0)
    pdf.cell(80, 5, "", border=0, new_x=XPos.LMARGIN, new_y=YPos.NEXT, align="C", markdown=True)

    for index, item in enumerate(front_page_data):
        if index > 0:
            pdf.cell(105, 4, '', border=0)
            pdf.cell(80, 4, "", border=0, new_x=XPos.LMARGIN, new_y=YPos.NEXT, align="C", markdown=True)

        pdf.set_text_color(0, 0, 0)
        pdf.set_font('helvetica', size=9)
        pdf.cell(105, 4, '', border=0)
        pdf.cell(80, 4, f"{item['title']}", border=0, new_x=XPos.LMARGIN, new_y=YPos.NEXT, align="C", markdown=True)
        pdf.set_text_color(128, 128, 128)
        pdf.set_font('helvetica', size=7)
        pdf.cell(105, 4, '', border=0)
        pdf.cell(80, 4, f"{item['value']}", border=0, new_x=XPos.LMARGIN, new_y=YPos.NEXT, align="C", markdown=True)

    pdf.cell(120, 5, '', border=0)
    pdf.cell(50, 5, "", border='B', new_x=XPos.LMARGIN, new_y=YPos.NEXT, align="C", markdown=True)
    pdf.cell(105, 5, '', border=0)
    pdf.cell(80, 5, "", border=0, new_x=XPos.LMARGIN, new_y=YPos.NEXT, align="C", markdown=True)

    # CREATE A GAP
    pdf.cell(105, 7, '', border=0)
    pdf.cell(80, 7, "", border=0, new_x=XPos.LMARGIN, new_y=YPos.NEXT, align="C", markdown=True)
    pdf.set_text_color(128, 128, 128)
    pdf.set_font('helvetica', size=6)
    pdf.cell(105, 4, '', border=0)

    pdf.multi_cell(80, 4,
                   "**Address:** Office 2, 1st Floor, 251 Durban Road, Bellville, 7535, Cape Town, South Africa \n"
                   "**Reg:** 2004/013672/07 | **VAT:** 4320223722\n"
                   "**Tel:** 021 919 9944\n**Email:** invest@opportunity.co.za\n**Website:** www.opportunity.co.za",
                   new_x=XPos.LMARGIN, new_y=YPos.NEXT, markdown=True,
                   align="C")

    # CREATE A GAP
    pdf.cell(105, 7, '', border=0)
    pdf.cell(80, 7, "", border=0, new_x=XPos.LMARGIN, new_y=YPos.NEXT, align="C", markdown=True)
    # make the font grey
    pdf.set_text_color(128, 128, 128)
    pdf.set_font('helvetica', size=6)
    pdf.cell(105, 4, '', border=0)

    pdf.multi_cell(80, 4,
                   "**Investment Disclaimer:** This is a private placement. Different types of investments involve "
                   "varying degrees of risk, and there can be no assurance that any specific investment will either "
                   "be suitable or profitable for a Lender or prospective Lender's investment portfolio. No investors "
                   "or prospective investors should assume that any information presented and/or made available by "
                   "Opportunity Private Capital or its Associates is a substitute for personalized individual advice "
                   "from an advisor or any other investment professional. No guarantees as to the success of the "
                   "investment or the projected return are offered. They have undertaken to present as much factual "
                   "information as is available. Every precaution has been taken to offer sufficient security for the "
                   "investment monies given by investors. Neither Opportunity Private Capital, nor the Borrower has "
                   "been registered as a regulated collective investment scheme pursuant to the Collective Investment "
                   "Schemes Control Act 45 of 2002 (\"CISCA\") and neither Opportunity Private Capital, "
                   "nor the Borrower, nor NSST, nor the Attorneys is licensed under CISCA. Note that this investment "
                   "product presented by Opportunity Private Capital is not a financial product that falls within the "
                   "scope of any category or sub-category regulated by the FSCA. The information contained in this "
                   "document does not constitute a financial service as defined in the Financial Advisory and "
                   "Intermediary Services Act nor is it intended to solicit investment or promote a financial product "
                   "in any way. Opportunity Private Capital do not provide investment, tax, legal or accounting "
                   "advice. This material has been prepared for informational purposes only. You should consult your "
                   "own investment, tax, legal and accounting advisors before engaging in any transaction.",
                   new_x=XPos.LMARGIN, new_y=YPos.NEXT, markdown=True,
                   align="C")

    # FRONT PAGE ENDS HERE
    # START OF THE INVESTMENT PAGE

    pdf.add_page()

    pdf.set_font('helvetica', size=8)
    # pdf.set_text_color(0, 0, 0)
    pdf.set_text_color(255, 255, 255)

    # set cell background color to black
    pdf.set_fill_color(0, 0, 0)

    pdf.cell(18, 5, 'Code', border=1, align="C", fill=True)
    pdf.cell(1, 5, '', border=0, align="C", fill=False)
    pdf.cell(18, 5, 'Days', border=1, align="C", fill=True)
    pdf.cell(1, 5, '', border=0, align="C", fill=False)

    pdf.cell(22, 5, 'Date', border=1, align="C", fill=True)
    pdf.cell(1, 5, '', border=0, align="C", fill=False)

    pdf.cell(22, 5, 'Interest %', border=1, align="C", fill=True)
    pdf.cell(1, 5, '', border=0, align="C", fill=False)

    pdf.cell(22, 5, 'Investment', border=1, align="C", fill=True)
    pdf.cell(1, 5, '', border=0, align="C", fill=False)

    pdf.cell(22, 5, 'Interest', border=1, align="C", fill=True)
    pdf.cell(1, 5, '', border=0, align="C", fill=False)

    pdf.cell(27, 5, 'Cumulative Interest', border=1, align="C", fill=True)
    pdf.cell(1, 5, '', border=0, align="C", fill=False)

    pdf.cell(22, 5, 'Balance', border=1, align="C", fill=True, new_x=XPos.LMARGIN, new_y=YPos.NEXT)

    for item in data1:
        if item['investment_amount'] != "":
            item['investment_amount'] = "R " + str(item['investment_amount'])
        if item['interest'] != "" and item['interest'] != 0:
            item['interest'] = "R " + str(item['interest'])
        else:
            item['interest'] = ""
        if item['cumulative_interest'] != "":
            item['cumulative_interest'] = "R " + str(item['cumulative_interest'])
        if item['investment_amount'] == "" and item['cumulative_interest'] == "":
            item['cumulative_interest'] = "R " + item['interest']
        if item['balance'] != "":
            item['balance'] = f"R {str(item['balance'])}"
         # if item['investment_amount'] != "" then item['effective_date'] = item['effective_date'] else item['effective_date'] = the last day of the month in item['effective_date'] as a datetime and formated as a string in the format of YYYY-MM-DD
        if item['investment_amount'] != "":
            # item['effective_date'] = the end of the current month of item['effective_date'] as a datetime and formated as a string in the format of YYYY-MM-DD

            item['interest_until'] = datetime.strptime(item['interest_until'], '%Y-%m-%d').strftime('%Y-%m-%d')





        pdf.set_text_color(128, 128, 128)

        pdf.cell(18, 10, item['opportunity_code'], border='B', align="C", fill=False)
        pdf.cell(1, 10, '', border='B', align="C", fill=False)
        pdf.cell(18, 10, str(item['days']), border='B', align="C", fill=False)
        pdf.cell(1, 10, '', border='B', align="C", fill=False)

        pdf.cell(22, 10, item['interest_until'], border='B', align="C", fill=False)
        pdf.cell(1, 10, '', border='B', align="C", fill=False)

        pdf.cell(22, 10, item['interest_rate'], border='B', align="C", fill=False)
        pdf.cell(1, 10, '', border='B', align="C", fill=False)

        pdf.cell(22, 10, item['investment_amount'], border='B', align="R", fill=False)
        pdf.cell(1, 10, '', border='B', align="C", fill=False)

        pdf.cell(22, 10, item['interest'], border='B', align="R", fill=False)
        pdf.cell(1, 10, '', border='B', align="C", fill=False)

        pdf.cell(27, 10, item['cumulative_interest'], border='B', align="R", fill=False)
        pdf.cell(1, 10, '', border='B', align="C", fill=False)

        pdf.cell(22, 10, item['balance'], border='B', align="R", fill=False, new_x=XPos.LMARGIN, new_y=YPos.NEXT)

    # INVESTMENT PAGE ENDS HERE
    # START IF DISCLAIMER PAGE
    # pdf.add_page()

    pdf.cell(105, 7, '', border=0)
    pdf.cell(80, 7, "", border=0, new_x=XPos.LMARGIN, new_y=YPos.NEXT, align="C", markdown=True)
    pdf.cell(105, 7, '', border=0)
    pdf.cell(80, 7, "", border=0, new_x=XPos.LMARGIN, new_y=YPos.NEXT, align="C", markdown=True)
    pdf.cell(105, 7, '', border=0)
    pdf.cell(80, 7, "", border=0, new_x=XPos.LMARGIN, new_y=YPos.NEXT, align="C", markdown=True)

    if statement_type == 1:
        pdf.set_font('helvetica', size=8)
        pdf.set_text_color(0, 0, 0)

        pdf.set_fill_color(240, 240, 240)

        pdf.cell(0, 3, "", border=0, align="C", fill=True, new_x=XPos.LMARGIN, new_y=YPos.NEXT)

        pdf.multi_cell(0, 4,
                       "**Earning of interest:** Interest calculations and accruals are reflected on the Lender "
                       "statements for informative purposes with interest on loans deemed earned on exit or repayment "
                       "date only. For the avoidance of any doubt, the Lender shall not earn any Interest until "
                       "Repayment"
                       "Date and the Interest shall only vest in the Lender on the Repayment Date.",
                       new_x=XPos.LMARGIN, new_y=YPos.NEXT, markdown=True,
                       align="L", fill=True)

        pdf.cell(0, 3, "", border=0, align="C", fill=True, new_x=XPos.LMARGIN, new_y=YPos.NEXT)

        pdf.set_fill_color(0, 0, 0)
        pdf.cell(0, 10, "Opportunity Private Capital - All rights reserved", border=0, align="C", fill=False,
                 new_x=XPos.LMARGIN, new_y=YPos.NEXT)

    else:

        pdf.set_font('helvetica', size=8)
        pdf.set_text_color(0, 0, 0)

        pdf.multi_cell(0, 4, f"**The lender represented here in by _ _ _ _ _ _ _ _ _ _ _ _being duly authorized to, "
                             f"accepts this lender statement as correct and final, and accepts the payment of the "
                             f"amount of "
                             f"{data1[len(data1) - 1]['balance']} as full and final settlement of any amounts due , "
                             f"arising from the loan agreement entered between parties with no party having any claim "
                             f"against the other from this day forth.**",
                       new_x=XPos.LMARGIN, new_y=YPos.NEXT, markdown=True,
                       align="L", fill=False)

        pdf.set_font('helvetica', size=6)
        pdf.cell(50, 20, "", border='B', align="C", fill=False)
        pdf.cell(50, 20, "", border=0, align="C", fill=False)
        pdf.cell(50, 20, "", border='B', align="C", fill=False, new_x=XPos.LMARGIN, new_y=YPos.NEXT)

        pdf.cell(50, 4, "LENDER SIGNATURE", border=0, align="L", fill=False)
        pdf.cell(50, 4, "", border=0, align="L", fill=False)
        pdf.cell(50, 4, "DATE", border=0, align="L", fill=False, new_x=XPos.LMARGIN, new_y=YPos.NEXT)

        pdf.cell(50, 20, "", border='B', align="C", fill=False)
        pdf.cell(50, 20, "", border=0, align="C", fill=False)
        pdf.cell(50, 20, "", border='B', align="C", fill=False, new_x=XPos.LMARGIN, new_y=YPos.NEXT)

        pdf.cell(50, 4, "MD van Rooyen", border=0, align="L", fill=False)
        pdf.cell(50, 4, "", border=0, align="L", fill=False)
        pdf.cell(50, 4, "DATE", border=0, align="L", fill=False, new_x=XPos.LMARGIN, new_y=YPos.NEXT)

        pdf.cell(50, 4, "NSST Trustee", border=0, align="L", fill=False, new_x=XPos.LMARGIN, new_y=YPos.NEXT)

    # END IF DISCLAIMER PAGE

    pdf.output(f'portal_statements/{filename}')

    return filename
