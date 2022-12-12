from fpdf import FPDF, XPos, YPos
import loan_agreement_files.contract as c
import loan_agreement_files.contents as cont
import loan_agreement_files.lender as l


class PDF(FPDF):
    def footer(self):
        self.set_y(-15)
        self.set_font('helvetica', 'I', 6)
        self.cell(0, 10, f"Page {self.page_no()}/{{nb}}", align="C")


# print(c.contract)
def print_investor_loan_agreement(lender, lender2, nsst, project, linked_unit, investment_amount, investment_interest_rate,
                                  investor_id, investor_id2):

    pdf = PDF('P', 'mm', 'A4')
    pdf.set_auto_page_break(auto=True, margin=15)
    pdf.add_page()
    pdf.set_font('helvetica', '', 10)
    pdf.set_fill_color(211, 211, 211)
    pdf.set_font('helvetica', '', 12)

    pdf.cell(0, 5, "**DEVELOPMENT INVESTMENT LOAN AGREEMENT**", new_x=XPos.LMARGIN, new_y=YPos.NEXT, markdown=True,
             align="C")
    pdf.set_font('helvetica', '', 10)
    pdf.cell(0, 5, "**Construction Phase**", new_x=XPos.LMARGIN, new_y=YPos.NEXT, markdown=True,
             align="C")
    pdf.cell(0, 5, "", new_x=XPos.LMARGIN, new_y=YPos.NEXT)
    pdf.cell(0, 5, "", new_x=XPos.LMARGIN, new_y=YPos.NEXT)
    pdf.cell(0, 5, "", new_x=XPos.LMARGIN, new_y=YPos.NEXT)
    pdf.cell(30, 5, "Project  ", align="R", markdown=True)
    pdf.cell(50, 5, f"{project}", align="L", markdown=True, border=True)
    pdf.cell(30, 5, "Linked Unit  ", align="R", markdown=True)
    pdf.cell(50, 5, f"{linked_unit}", align="L", markdown=True, border=True, new_x=XPos.LMARGIN, new_y=YPos.NEXT)
    pdf.cell(0, 5, "", new_x=XPos.LMARGIN, new_y=YPos.NEXT)
    pdf.cell(30, 5, "between  ", new_x=XPos.LMARGIN, new_y=YPos.NEXT, align="R")
    pdf.cell(0, 5, "", new_x=XPos.LMARGIN, new_y=YPos.NEXT)
    pdf.cell(0, 5, "", new_x=XPos.LMARGIN, new_y=YPos.NEXT)
    pdf.cell(30, 5, "", align="R", markdown=True)
    pdf.cell(50, 5, "**HERON PROJECTS PROPRIETARY LIMITED**", align="L", markdown=True, border=False,
             new_x=XPos.LMARGIN, new_y=YPos.NEXT)

    pdf.cell(30, 5, "", align="R", markdown=True)
    pdf.cell(50, 5, "Registration Number **2020/495056/07**", align="L", markdown=True, border=False,
             new_x=XPos.LMARGIN, new_y=YPos.NEXT)
    pdf.cell(30, 5, "", align="R", markdown=True)
    pdf.cell(50, 5, "Herein represented by **Charles Nixon Morgan**", align="L", markdown=True, border=False,
             new_x=XPos.LMARGIN, new_y=YPos.NEXT)
    pdf.cell(30, 5, "", align="R", markdown=True)
    pdf.cell(50, 5, "Duly authorised in terms of a resolution", align="L", markdown=True, border=False,
             new_x=XPos.LMARGIN, new_y=YPos.NEXT)
    pdf.cell(0, 5, "", new_x=XPos.LMARGIN, new_y=YPos.NEXT)
    pdf.cell(0, 5, "", new_x=XPos.LMARGIN, new_y=YPos.NEXT)
    pdf.cell(30, 5, "", align="R", markdown=True)
    pdf.cell(50, 5, "(hereinafter referred to as 'the Borrower')", align="L", markdown=True, border=False,
             new_x=XPos.LMARGIN, new_y=YPos.NEXT)
    pdf.cell(0, 5, "", new_x=XPos.LMARGIN, new_y=YPos.NEXT)
    pdf.cell(0, 5, "", new_x=XPos.LMARGIN, new_y=YPos.NEXT)
    pdf.cell(30, 5, "", align="R", markdown=True)
    pdf.cell(50, 5, "and", align="L", markdown=True, border=False,
             new_x=XPos.LMARGIN, new_y=YPos.NEXT)

    pdf.cell(0, 5, "", new_x=XPos.LMARGIN, new_y=YPos.NEXT)
    pdf.cell(0, 5, "", new_x=XPos.LMARGIN, new_y=YPos.NEXT)
    pdf.cell(30, 5, "", align="R", markdown=True)
    if not lender == "":
        pdf.cell(50, 5, f"**{lender}**", align="L", markdown=True, border=False,
                 new_x=XPos.LMARGIN, new_y=YPos.NEXT)
    else:
        pdf.cell(50, 5, "__________________________________", align="L", markdown=True, border=False,
                 new_x=XPos.LMARGIN, new_y=YPos.NEXT)
    pdf.cell(30, 5, "", align="R", markdown=True)
    if investor_id == "":
        pdf.cell(50, 5, "Identity/Registration Number _______________________", align="L", markdown=True, border=False,
                 new_x=XPos.LMARGIN, new_y=YPos.NEXT)
    else:
        pdf.cell(50, 5, f"Identity/Registration Number {investor_id}", align="L", markdown=True, border=False,
                 new_x=XPos.LMARGIN, new_y=YPos.NEXT)

    pdf.cell(30, 5, "", align="R", markdown=True)
    pdf.cell(50, 5, "Herein represented by ___________________ (delete if not applicable)", align="L", markdown=True,
             border=False,
             new_x=XPos.LMARGIN, new_y=YPos.NEXT)
    pdf.cell(30, 5, "", align="R", markdown=True)
    pdf.cell(50, 5, "Duly authorised in terms of a resolution (delete if not applicable)", align="L", markdown=True,
             border=False,
             new_x=XPos.LMARGIN, new_y=YPos.NEXT)
    # pdf.cell(0, 5, "", new_x=XPos.LMARGIN, new_y=YPos.NEXT)
    # pdf.cell(0, 5, "", new_x=XPos.LMARGIN, new_y=YPos.NEXT)
    pdf.cell(30, 5, "", align="R", markdown=True)
    pdf.cell(50, 5, "(hereinafter referred to as 'the Lender No1')", align="L", markdown=True, border=False,
             new_x=XPos.LMARGIN, new_y=YPos.NEXT)
    pdf.cell(0, 5, "", new_x=XPos.LMARGIN, new_y=YPos.NEXT)
    pdf.cell(0, 5, "", new_x=XPos.LMARGIN, new_y=YPos.NEXT)
    pdf.cell(30, 5, "", align="R", markdown=True)
    pdf.cell(50, 5, "and", align="L", markdown=True, border=False,
             new_x=XPos.LMARGIN, new_y=YPos.NEXT)

    # pdf.cell(0, 5, "", new_x=XPos.LMARGIN, new_y=YPos.NEXT)
    # pdf.cell(0, 5, "", new_x=XPos.LMARGIN, new_y=YPos.NEXT)
    # pdf.cell(30, 5, "", align="R", markdown=True)
    # pdf.cell(50, 5, "and", align="L", markdown=True, border=False,
    #          new_x=XPos.LMARGIN, new_y=YPos.NEXT)

    pdf.cell(0, 5, "", new_x=XPos.LMARGIN, new_y=YPos.NEXT)
    pdf.cell(0, 5, "", new_x=XPos.LMARGIN, new_y=YPos.NEXT)
    pdf.cell(30, 5, "", align="R", markdown=True)
    if not lender2 == "":
        pdf.cell(50, 5, f"**{lender2}**", align="L", markdown=True, border=False,
                 new_x=XPos.LMARGIN, new_y=YPos.NEXT)
    else:
        pdf.cell(50, 5, "__________________________________", align="L", markdown=True, border=False,
                 new_x=XPos.LMARGIN, new_y=YPos.NEXT)
    pdf.cell(30, 5, "", align="R", markdown=True)
    if investor_id2 == "":
        pdf.cell(50, 5, "Identity/Registration Number _______________________", align="L", markdown=True, border=False,
                 new_x=XPos.LMARGIN, new_y=YPos.NEXT)
    else:
        pdf.cell(50, 5, f"Identity/Registration Number {investor_id2}", align="L", markdown=True, border=False,
                 new_x=XPos.LMARGIN, new_y=YPos.NEXT)

    pdf.cell(30, 5, "", align="R", markdown=True)
    pdf.cell(50, 5, "Herein represented by ___________________ (delete if not applicable)", align="L", markdown=True,
             border=False,
             new_x=XPos.LMARGIN, new_y=YPos.NEXT)
    pdf.cell(30, 5, "", align="R", markdown=True)
    pdf.cell(50, 5, "Duly authorised in terms of a resolution (delete if not applicable)", align="L", markdown=True,
             border=False,
             new_x=XPos.LMARGIN, new_y=YPos.NEXT)
    # pdf.cell(0, 5, "", new_x=XPos.LMARGIN, new_y=YPos.NEXT)
    # pdf.cell(0, 5, "", new_x=XPos.LMARGIN, new_y=YPos.NEXT)
    pdf.cell(30, 5, "", align="R", markdown=True)
    pdf.cell(50, 5, "(hereinafter referred to as 'the Lender No2')", align="L", markdown=True, border=False,
             new_x=XPos.LMARGIN, new_y=YPos.NEXT)
    pdf.cell(0, 5, "", new_x=XPos.LMARGIN, new_y=YPos.NEXT)
    pdf.cell(0, 5, "", new_x=XPos.LMARGIN, new_y=YPos.NEXT)
    pdf.cell(30, 5, "", align="R", markdown=True)
    pdf.cell(50, 5, "and", align="L", markdown=True, border=False,
             new_x=XPos.LMARGIN, new_y=YPos.NEXT)

    pdf.cell(0, 5, "", new_x=XPos.LMARGIN, new_y=YPos.NEXT)
    pdf.cell(0, 5, "", new_x=XPos.LMARGIN, new_y=YPos.NEXT)
    pdf.cell(30, 5, "", align="R", markdown=True)
    pdf.cell(50, 5, "**NOBLE SHIELD 2 SECURITY TRUST**", align="L", markdown=True, border=False,
             new_x=XPos.LMARGIN, new_y=YPos.NEXT)
    pdf.cell(30, 5, "", align="R", markdown=True)
    pdf.cell(50, 5, "Registration Number **IT000426/2021(c)**", align="L", markdown=True, border=False,
             new_x=XPos.LMARGIN, new_y=YPos.NEXT)
    pdf.cell(30, 5, "", align="R", markdown=True)
    pdf.cell(50, 5, "Herein represented by **Martin Deon van Rooyen CA(SA), RA**", align="L", markdown=True,
             border=False,
             new_x=XPos.LMARGIN, new_y=YPos.NEXT)
    pdf.cell(30, 5, "", align="R", markdown=True)
    pdf.cell(50, 5, "Duly authorised in terms of a resolution", align="L", markdown=True,
             border=False,
             new_x=XPos.LMARGIN, new_y=YPos.NEXT)
    pdf.cell(0, 5, "", new_x=XPos.LMARGIN, new_y=YPos.NEXT)
    pdf.cell(0, 5, "", new_x=XPos.LMARGIN, new_y=YPos.NEXT)
    pdf.cell(30, 5, "", align="R", markdown=True)
    pdf.cell(50, 5, "(hereinafter referred to as 'NSST')", align="L", markdown=True, border=False,
             new_x=XPos.LMARGIN, new_y=YPos.NEXT)

    pdf.add_page()

    pdf.cell(190, 10, "**LENDER'S INFORMATION SCHEDULE**", align="C", markdown=True, fill=True, border=True,
             new_x=XPos.LMARGIN, new_y=YPos.NEXT)
    for i in l.lender_info:
        pdf.cell(80, 7, f"{i['Label']}", align="L", markdown=True, fill=True, border=i['border1'])
        pdf.cell(110, 7, f"{i['text']}", align="L", markdown=True, fill=False, border=i['border2'], new_x=XPos.LMARGIN,
                 new_y=YPos.NEXT)

    pdf.cell(0, 5, "", new_x=XPos.LMARGIN, new_y=YPos.NEXT)
    pdf.cell(0, 5, "", new_x=XPos.LMARGIN, new_y=YPos.NEXT)
    pdf.cell(0, 5, "", new_x=XPos.LMARGIN, new_y=YPos.NEXT)
    pdf.cell(190, 10, "**INVESTMENT INFORMATION SCHEDULE**", align="C", markdown=True, fill=True, border=True,
             new_x=XPos.LMARGIN, new_y=YPos.NEXT)
    pdf.cell(80, 7, "Interest rate (%) per annum", align="L", markdown=True, fill=True, border=True)
    pdf.cell(110, 7, f"{investment_interest_rate} %", align="L", markdown=True, fill=False, border=True,
             new_x=XPos.LMARGIN,
             new_y=YPos.NEXT)
    pdf.cell(80, 7, "Loan Amount (R)", align="L", markdown=True, fill=True, border=True)
    pdf.cell(110, 7, f"{investment_amount}", align="L", markdown=True, fill=False, border=True, new_x=XPos.LMARGIN,
             new_y=YPos.NEXT)
    pdf.cell(80, 7, "Property", align="L", markdown=True, fill=True, border=True)
    pdf.cell(110, 7, f"{project}", align="L", markdown=True, fill=False, border=True,
             new_x=XPos.LMARGIN,
             new_y=YPos.NEXT)
    pdf.cell(80, 7, "Linked Unit", align="L", markdown=True, fill=True, border=True)
    pdf.cell(110, 7, f"{linked_unit}", align="L", markdown=True, fill=False, border=True,
             new_x=XPos.LMARGIN,
             new_y=YPos.NEXT)

    pdf.add_page()
    pdf.cell(0, 5, "", new_x=XPos.LMARGIN, new_y=YPos.NEXT)
    pdf.cell(0, 5, "**TABLE OF CONTENTS**", new_x=XPos.LMARGIN, new_y=YPos.NEXT, markdown=True, align="C")
    pdf.cell(0, 5, "", new_x=XPos.LMARGIN, new_y=YPos.NEXT)
    for i in cont.contentsList:
        pdf.cell(50, 5, f"**{i['number']}** ", align="R", markdown=True)
        pdf.cell(100, 5, f"**{i['text']}**", markdown=True)
        pdf.cell(10, 5, f"**{i['page']}**", new_x=XPos.LMARGIN, new_y=YPos.NEXT, markdown=True, align="R")
        # pdf.cell(0, 5, "", new_x=XPos.LMARGIN, new_y=YPos.NEXT)

    pdf.add_page()
    for line in c.contract:
        if c.contract[line]["indent"] == 1:
            pdf.cell(0, 5, "", new_x=XPos.LMARGIN, new_y=YPos.NEXT)
            pdf.cell(20, 5, f'{line}', markdown=True, align="R")
            pdf.multi_cell(170, 5, c.contract[line]["text"], new_x=XPos.LMARGIN, new_y=YPos.NEXT, border=False,
                           markdown=True)
        elif c.contract[line]["indent"] == 0:
            pdf.cell(0, 5, "", new_x=XPos.LMARGIN, new_y=YPos.NEXT)
            pdf.cell(40, 5, c.contract[line]["text"], new_x=XPos.LMARGIN, new_y=YPos.NEXT, border=False, markdown=True)
        elif c.contract[line]["indent"] == -1:
            pdf.cell(0, 5, c.contract[line]["text"],
                     new_x=XPos.LMARGIN,
                     new_y=YPos.NEXT, border=False,
                     markdown=True)
        elif c.contract[line]["indent"] == 2:
            pdf.cell(0, 5, "", new_x=XPos.LMARGIN, new_y=YPos.NEXT)
            pdf.cell(30, 5, f'{line}', markdown=True, align="R")
            pdf.multi_cell(160, 5, c.contract[line]["text"], new_x=XPos.LMARGIN, new_y=YPos.NEXT, border=False,
                           markdown=True)
        elif c.contract[line]["indent"] == 3:
            pdf.cell(0, 5, "", new_x=XPos.LMARGIN, new_y=YPos.NEXT)
            pdf.cell(40, 5, f'{line}', markdown=True, align="R")
            pdf.multi_cell(150, 5, c.contract[line]["text"], new_x=XPos.LMARGIN, new_y=YPos.NEXT, border=False,
                           markdown=True)
        elif c.contract[line]["indent"] == 999:
            # if c.contract[line]["first"] == "Yes":
            #     pdf.cell(0, 5, "", new_x=XPos.LMARGIN, new_y=YPos.NEXT)
            pdf.cell(40, 5, "")
            pdf.cell(40, 5, c.contract[line]["text"], markdown=True, align="L", fill=True, border=True)
            pdf.multi_cell(50, 5, c.contract[line]["text1"], new_x=XPos.LMARGIN, new_y=YPos.NEXT, border=True,
                           markdown=True)
        elif c.contract[line]["indent"] == 111:
            pdf.cell(0, 5, "", new_x=XPos.LMARGIN, new_y=YPos.NEXT)
            pdf.cell(10, 5, "")
            pdf.multi_cell(180, 5, c.contract[line]["text"], new_x=XPos.LMARGIN, new_y=YPos.NEXT, border=False,
                           markdown=True)
        elif c.contract[line]["indent"] == 333:
            pdf.cell(0, 5, "", new_x=XPos.LMARGIN, new_y=YPos.NEXT)
            pdf.set_font('helvetica', '', 8)
            pdf.cell(20, 3, "")
            pdf.multi_cell(160, 3, c.contract[line]["text"], new_x=XPos.LMARGIN, new_y=YPos.NEXT, border=False,
                           markdown=True)
            pdf.cell(10, 3, "")
            pdf.set_font('helvetica', '', 10)

    pdf.cell(0, 5, "", new_x=XPos.LMARGIN, new_y=YPos.NEXT)
    pdf.cell(0, 5, "**SIGNATURE PAGE OVERLEAF**", new_x=XPos.LMARGIN, new_y=YPos.NEXT, markdown=True, align="C")
    pdf.add_page()
    pdf.cell(0, 5, "", new_x=XPos.LMARGIN, new_y=YPos.NEXT)
    pdf.cell(0, 5, "", new_x=XPos.LMARGIN, new_y=YPos.NEXT)
    pdf.cell(0, 5, "", new_x=XPos.LMARGIN, new_y=YPos.NEXT)
    pdf.cell(0, 5, "", new_x=XPos.LMARGIN, new_y=YPos.NEXT)
    pdf.cell(0, 5, "SIGNED AT_________________________ THIS ______DAY OF ________________202___", new_x=XPos.LMARGIN,
             new_y=YPos.NEXT, markdown=True, align="C")
    pdf.cell(0, 5, "", new_x=XPos.LMARGIN, new_y=YPos.NEXT)
    pdf.cell(30, 5, "")
    pdf.cell(160, 5, "IN THE PRESENCE OF THE FOLLOWING WITNESSES", new_x=XPos.LMARGIN, new_y=YPos.NEXT)
    pdf.cell(0, 5, "", new_x=XPos.LMARGIN, new_y=YPos.NEXT)
    pdf.cell(0, 5, "", new_x=XPos.LMARGIN, new_y=YPos.NEXT)
    pdf.cell(0, 5, "", new_x=XPos.LMARGIN, new_y=YPos.NEXT)
    pdf.cell(0, 5, "", new_x=XPos.LMARGIN, new_y=YPos.NEXT)
    pdf.cell(0, 5, "", new_x=XPos.LMARGIN, new_y=YPos.NEXT)
    pdf.cell(95, 5, " _______________________", align="C")
    pdf.cell(95, 5, " _______________________", new_x=XPos.LMARGIN, new_y=YPos.NEXT, align="C")
    pdf.cell(95, 5, " WITNESS", align="C")
    pdf.cell(95, 5, "  FOR: THE BORROWER", new_x=XPos.LMARGIN, new_y=YPos.NEXT, align="C")
    pdf.cell(123, 5, "")
    pdf.cell(67, 5, "NAME: **Charles Nixon Morgan**", new_x=XPos.LMARGIN, new_y=YPos.NEXT, markdown=True)
    pdf.cell(123, 5, "")
    pdf.set_font('helvetica', '', 8)
    pdf.cell(67, 5, "(who warrants that he is duly authorised)", new_x=XPos.LMARGIN, new_y=YPos.NEXT, markdown=True)
    pdf.set_font('helvetica', '', 10)
    pdf.cell(0, 5, "", new_x=XPos.LMARGIN, new_y=YPos.NEXT)
    pdf.cell(0, 5, "", new_x=XPos.LMARGIN, new_y=YPos.NEXT)
    pdf.cell(0, 5, "SIGNED AT_________________________ THIS ______DAY OF ________________202___", new_x=XPos.LMARGIN,
             new_y=YPos.NEXT, markdown=True, align="C")
    pdf.cell(0, 5, "", new_x=XPos.LMARGIN, new_y=YPos.NEXT)
    pdf.cell(30, 5, "")
    pdf.cell(160, 5, "IN THE PRESENCE OF THE FOLLOWING WITNESSES", new_x=XPos.LMARGIN, new_y=YPos.NEXT)
    pdf.cell(0, 5, "", new_x=XPos.LMARGIN, new_y=YPos.NEXT)
    pdf.cell(0, 5, "", new_x=XPos.LMARGIN, new_y=YPos.NEXT)
    pdf.cell(0, 5, "", new_x=XPos.LMARGIN, new_y=YPos.NEXT)
    pdf.cell(0, 5, "", new_x=XPos.LMARGIN, new_y=YPos.NEXT)
    pdf.cell(0, 5, "", new_x=XPos.LMARGIN, new_y=YPos.NEXT)
    pdf.cell(95, 5, " _______________________", align="C")
    pdf.cell(95, 5, " _______________________", new_x=XPos.LMARGIN, new_y=YPos.NEXT, align="C")
    pdf.cell(95, 5, " WITNESS", align="C")
    pdf.cell(95, 5, "  FOR: THE LENDER", new_x=XPos.LMARGIN, new_y=YPos.NEXT, align="C")
    pdf.cell(123, 5, "")
    pdf.cell(67, 5, f"NAME: **{lender}**", new_x=XPos.LMARGIN, new_y=YPos.NEXT, markdown=True)
    pdf.cell(123, 5, "")
    pdf.set_font('helvetica', '', 8)
    pdf.cell(67, 5, "(who warrants that he is duly authorised)", new_x=XPos.LMARGIN, new_y=YPos.NEXT, markdown=True)
    pdf.set_font('helvetica', '', 10)
    pdf.cell(0, 5, "", new_x=XPos.LMARGIN, new_y=YPos.NEXT)
    pdf.cell(0, 5, "", new_x=XPos.LMARGIN, new_y=YPos.NEXT)
    pdf.cell(0, 5, "SIGNED AT_________________________ THIS ______DAY OF ________________202___", new_x=XPos.LMARGIN,
             new_y=YPos.NEXT, markdown=True, align="C")
    pdf.cell(0, 5, "", new_x=XPos.LMARGIN, new_y=YPos.NEXT)
    pdf.cell(30, 5, "")
    pdf.cell(160, 5, "IN THE PRESENCE OF THE FOLLOWING WITNESSES", new_x=XPos.LMARGIN, new_y=YPos.NEXT)
    pdf.cell(0, 5, "", new_x=XPos.LMARGIN, new_y=YPos.NEXT)
    pdf.cell(0, 5, "", new_x=XPos.LMARGIN, new_y=YPos.NEXT)
    pdf.cell(0, 5, "", new_x=XPos.LMARGIN, new_y=YPos.NEXT)
    pdf.cell(0, 5, "", new_x=XPos.LMARGIN, new_y=YPos.NEXT)
    pdf.cell(0, 5, "", new_x=XPos.LMARGIN, new_y=YPos.NEXT)
    pdf.cell(95, 5, " _______________________", align="C")
    pdf.cell(95, 5, " _______________________", new_x=XPos.LMARGIN, new_y=YPos.NEXT, align="C")
    pdf.cell(95, 5, " WITNESS", align="C")
    pdf.cell(95, 5, "  FOR: NSST", new_x=XPos.LMARGIN, new_y=YPos.NEXT, align="C")
    pdf.cell(123, 5, "")
    pdf.cell(67, 5, f"NAME: **{nsst}**", new_x=XPos.LMARGIN, new_y=YPos.NEXT, markdown=True)
    pdf.cell(123, 5, "")
    pdf.set_font('helvetica', '', 8)
    pdf.cell(67, 5, "(who warrants that he is duly authorised)", new_x=XPos.LMARGIN, new_y=YPos.NEXT, markdown=True)
    pdf.set_font('helvetica', '', 10)

    pdf.output("LoanAgreement.pdf")

    return "LoanAgreement.pdf"
