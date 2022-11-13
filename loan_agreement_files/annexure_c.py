from fpdf import FPDF, XPos, YPos
# from loan_agreement_files.loan_agreement import lender, nsst


class PDF(FPDF):
    def footer(self):
        self.set_y(-15)
        self.set_font('helvetica', 'I', 6)
        self.cell(0, 10, f"Page {self.page_no()}/{{nb}}", align="C")


def print_annexure_c(lender, nsst):
    pdf = PDF('P', 'mm', 'A4')
    pdf.set_auto_page_break(auto=True, margin=15)
    pdf.add_page()
    pdf.set_font('helvetica', '', 10)
    pdf.set_fill_color(211, 211, 211)
    pdf.set_font('helvetica', '', 12)

    pdf.cell(0, 5, "", new_x=XPos.LMARGIN, new_y=YPos.NEXT)
    pdf.cell(0, 5, "", new_x=XPos.LMARGIN, new_y=YPos.NEXT)
    pdf.cell(150, 5, "")
    pdf.cell(30, 5, "**ANNEXURE C**", new_x=XPos.LMARGIN, new_y=YPos.NEXT, align="R", markdown=True, border="B")
    pdf.set_font('helvetica', '', 10)
    pdf.cell(0, 5, "", new_x=XPos.LMARGIN, new_y=YPos.NEXT)
    pdf.cell(0, 5, "", new_x=XPos.LMARGIN, new_y=YPos.NEXT)
    pdf.cell(50, 5, "**Instruction to Attorneys**", new_x=XPos.LMARGIN, new_y=YPos.NEXT, border="B", markdown=True)
    pdf.cell(0, 5, "", new_x=XPos.LMARGIN, new_y=YPos.NEXT)
    pdf.cell(0, 5, "", new_x=XPos.LMARGIN, new_y=YPos.NEXT)
    pdf.multi_cell(0, 5, "We, the undersigned, confirm that the funds paid into the Trust Account of the Attorneys ("
                         "LAÄS & SCHOLTZ INC) are held for the benefit of the Borrower, **HERON PROJECTS PROPRIETARY "
                         "LIMITED**, Registration Number **2020/495056/07**, but insofar as it may be necessary, "
                         "we consent "
                         "to the funds being paid to the Investment Account.", markdown=True)
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

    pdf.output("annexure_c.pdf")

    return "annexure_c.pdf"
