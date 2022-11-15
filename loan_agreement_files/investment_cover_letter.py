from fpdf import FPDF, XPos, YPos
# from loan_agreement_files.loan_agreement import lender


class PDF(FPDF):
    def footer(self):
        self.set_y(-15)
        self.set_font('helvetica', 'I', 6)
        self.cell(0, 10, f"Page {self.page_no()}/{{nb}}", align="C")


def print_investor_cover_letter(lender):
    pdf = PDF('P', 'mm', 'A4')
    pdf.set_auto_page_break(auto=True, margin=15)
    pdf.add_page()
    pdf.set_font('helvetica', '', 10)
    pdf.set_fill_color(211, 211, 211)
    pdf.set_font('helvetica', '', 12)

    pdf.image("loan_agreement_files/Latest header.jpg", 10, 10, 190)
    pdf.cell(0, 70, "", new_x=XPos.LMARGIN, new_y=YPos.NEXT)
    pdf.cell(0, 5, "**INVESTMENT DISCLAIMER**", new_x=XPos.LMARGIN, new_y=YPos.NEXT, markdown=True,
             align="C")
    pdf.cell(0, 5, "", new_x=XPos.LMARGIN, new_y=YPos.NEXT)
    pdf.cell(0, 5, "", new_x=XPos.LMARGIN, new_y=YPos.NEXT)
    if lender == "":
        pdf.cell(0, 5, "**LENDER NAME / ENTITY NAME: _________________________________________________**",
                 new_x=XPos.LMARGIN, new_y=YPos.NEXT, markdown=True,
                 align="C")
    else:
        pdf.cell(0, 5, f"**LENDER NAME / ENTITY NAME: {lender}**",
                 new_x=XPos.LMARGIN, new_y=YPos.NEXT, markdown=True,
                 align="C")

    pdf.cell(0, 5, "", new_x=XPos.LMARGIN, new_y=YPos.NEXT)
    pdf.cell(0, 5, "", new_x=XPos.LMARGIN, new_y=YPos.NEXT)

    pdf.multi_cell(0, 5,
                   "**Kindly take note of the terms stated hereunder and sign this document as acknowledgement hereof.**",
                   new_x=XPos.LMARGIN, new_y=YPos.NEXT, markdown=True,
                   align="C")

    pdf.cell(0, 5, "", new_x=XPos.LMARGIN, new_y=YPos.NEXT)
    pdf.cell(0, 5, "", new_x=XPos.LMARGIN, new_y=YPos.NEXT)
    pdf.cell(10, 5, "")
    # pdf.set_stretching(150)
    pdf.multi_cell(170, 5,
                   "This is a private placement. Different types of investments involve varying degrees of risk, "
                   "and there can be no assurance that any specific investment will either be suitable or profitable "
                   "for a Lender or prospective Lender's investment portfolio.  No investors or prospective investors "
                   "should assume that any information presented and/or made available by Opportunity Private Capital "
                   "or its Associates is a substitute for personalized individual advice from an advisor or any other "
                   "investment professional. No guarantees as to the success of the investment or the projected "
                   "return are offered. They have undertaken to present as much factual information as is available.  "
                   "Every precaution has been taken to offer sufficient security for the investment monies given by "
                   "investors. Neither Opportunity Private Capital, nor the Borrower has been registered as a "
                   "regulated collective investment scheme pursuant to the Collective Investment Schemes Control Act "
                   "45 of 2002 ('CISCA') and neither Opportunity Private Capital, nor the Borrower, nor NSST, "
                   "nor the Attorneys is licensed under CISCA. The information contained in this document does not "
                   "constitute a financial service as defined in the Financial Advisory and Intermediary Services Act "
                   "nor is it intended to solicit investment or promote a financial product in any way. Opportunity "
                   "Private Capital do not provide investment, tax, legal or accounting advice. This material has "
                   "been prepared for informational purposes only. You should consult your own investment, tax, "
                   "legal and accounting advisors before engaging in any transaction.",
                   new_x=XPos.LMARGIN, new_y=YPos.NEXT, markdown=True,
                   align="J")
    pdf.cell(10, 5, "")

    pdf.cell(0, 5, "", new_x=XPos.LMARGIN, new_y=YPos.NEXT)
    pdf.cell(0, 5, "", new_x=XPos.LMARGIN, new_y=YPos.NEXT)

    pdf.multi_cell(0, 5,
                   "**LENDER SIGN: ___________________________**",
                   new_x=XPos.LMARGIN, new_y=YPos.NEXT, markdown=True,
                   align="L")

    pdf.cell(0, 5, "", new_x=XPos.LMARGIN, new_y=YPos.NEXT)
    pdf.cell(0, 5, "", new_x=XPos.LMARGIN, new_y=YPos.NEXT)

    pdf.multi_cell(0, 5,
                   "**DATE: ___________________________**",
                   new_x=XPos.LMARGIN, new_y=YPos.NEXT, markdown=True,
                   align="L")

    pdf.image("loan_agreement_files/Footer July.jpg", 10, 277, 190)

    pdf.output("Investment_cover_letter.pdf")

    return "Investment_cover_letter.pdf"
