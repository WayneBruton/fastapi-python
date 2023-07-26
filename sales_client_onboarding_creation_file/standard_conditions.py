def create_standard_conditions(data):
    standard_conditions = [
        {
            "type": 1,
            "section": "1.",
            "text": "**PREAMBLE**",
            "initial": False,
        },
        {
            "type": 2,
            "section": "1.1",
            "text": "The Seller has agreed to sell, and the Purchaser has agreed to purchase the Property in C"
                    f" of the Information Schedule, to be established in the Sectional Title Scheme to be "
                    f"known as **\"{data['development'].upper()}\"** in terms of the Sectional Titles Act.",
            "initial": False,
        },
        {
            "type": 2,
            "section": "1.2",
            "text": "The Sale is subject to the fulfilment of the condition's precedent recorded "
                    f"in this Agreement.",
            "initial": False,
        },
        {
            "type": 1,
            "section": "2.",
            "text": "**INTERPRETATION**",
            "initial": False,
        },
        {
            "type": 2,
            "section": "2.1",
            "text": "In this Agreement, unless the context otherwise indicates:",
            "initial": False,
        },
        {
            "type": 3,
            "section": "2.1.1",
            "text": "**\"Architect\"** means any registered architect as may be appointed by the Seller "
                    f"from time to time.",
            "initial": False,
        },
        {
            "type": 3,
            "section": "2.1.2",
            "text": "**\"Beneficial Occupation\"** means the Property has water, power, sewerage, access and "
                    f"is thus liveable and ready for physical occupation.",
            "initial": False,
        },
        {
            "type": 3,
            "section": "2.1.3",
            "text": "**\"Body Corporate\"** means the controlling body of the Scheme as contemplated in terms"
                    f" of Section 36 of the Sectional Titles Act, which will come into existence with the"
                    f" transfer of the first Unit from the Seller to a Purchaser in this Scheme.",
            "initial": False,
        },
        {
            "type": 3,
            "section": "2.1.4",
            "text": "**\"Building\"** means the building/s to be constructed on the Land.",
            "initial": False,
        },
        {
            "type": 3,
            "section": "2.1.5",
            "text": "**\"Chief Ombud\"** means Chief Ombud as defined in Section 1 of the Community Schemes Ombud "
                    "Service Act, 2010.",
            "initial": False,
        },
        {
            "type": 3,
            "section": "2.1.6",
            "text": "**\"Completion Date\"** means the date upon which the building inspector employed by the local "
                    "authority or Architect issues an Occupation Certificate in respect of the Unit to the effect "
                    "that the Unit is fit for Beneficial Occupation, or the date of handover of the keys of the Unit "
                    "to the Purchaser, whichever date is earlier, subject to the provision that in the event of a "
                    "dispute, the Completion Date shall be certified as such by the Architect, whose decision as to "
                    "that date shall be final and binding on the parties.",
            "initial": False,
        },
        {
            "type": 3,
            "section": "2.1.7",
            "text": "**\"Common Property\"** means common property as defined in the Sectional Titles Act.",
            "initial": False,
        },
        {
            "type": 3,
            "section": "2.1.8",
            "text": "**\"Developer\"** means the party described as Seller in the Information Schedule.",
            "initial": False,
        },
        {
            "type": 3,
            "section": "2.1.9",
            "text": "**\"Development Period\"** means the period from the establishment of the Body Corporate to the "
                    "transfer of the last saleable sectional title unit in the Scheme or a period not exceeding "
                    "twenty years from date of establishment of the Body Corporate, whichever is the longest.",
            "initial": False,
        },
        {
            "type": 3,
            "section": "2.1.10",
            "text": "**\"Exclusive Use Areas\"** means such parts of the Common Property reserved for the exclusive "
                    "use and enjoyment of the registered owner for the time being of the Unit, in terms of Section "
                    "27A of the Sectional Titles Act which includes the Garden and Parking.",
            "initial": False,
        },
        {
            "type": 3,
            "section": "2.1.11",
            "text": "**\"FICA\"** means the Financial Intelligence Centre Act, Act 38 of 2001, as amended from time "
                    "to time.",
            "initial": False,
        },
        {
            "type": 3,
            "section": "2.1.12",
            "text": "**\"Happy Letter\"** means a formal document prepared in a format acceptable to the bank or "
                    "other recognised financial institution providing a bond to the Purchaser as provided for in "
                    "Clause 19 hereunder and signed by the Purchaser (or his Agent/Proxy) at the instance of the bank "
                    "or other recognised financial institution providing a bond, or in the case of a cash purchase, "
                    "a Happy Letter document provided by the Seller, to be signed by the Purchaser.",
            "initial": False,
        },
        {
            "type": 3,
            "section": "2.1.13",
            "text": "**\"Information Schedule\"** means the Information Schedule set out on pages 3, 4 and 5 hereof "
                    "which shall be deemed to be incorporated in this Agreement and shall be an integral part thereof.",
            "initial": False,
        },
        {
            "type": 3,
            "section": "2.1.14",
            "text": "**\"Land\"** means the land on which the Scheme is to be developed being Erf 41409 Kraaifontein, "
                    "(previously known as the Remainder of Portion 18 of the Farm Langeberg Number 311.",
            "initial": False,
        },
        {
            "type": 3,
            "section": "2.1.15",
            "text": "**\"Occupation Certificate\"** means a certificate issued by the City of Cape Town confirming "
                    "that the Unit is ready for Beneficial Occupation.",
            "initial": False,
        },
        {
            "type": 3,
            "section": "2.1.16",
            "text": "**\"Occupation Date\"** means the date on which the Seller hands over the keys of the Unit to "
                    "the Purchaser, or Transfer Date, whichever is the earliest.",
            "initial": False,

        },
        {
            "type": 3,
            "section": "2.1.17",
            "text": "**\"Prime Rate\"** means a rate of interest per annum which is equal to the published minimum "
                    "lending rate of interest per annum, compounded monthly in arrear, charged by ABSA Bank Limited "
                    "on the unsecured overdrawn current accounts of its most favoured corporate clients in the "
                    "private sector from time to time.  (In the case of a dispute as to the rate so payable, "
                    "the rate may be certified by any manager or assistant manager of any branch of the said bank, "
                    "whose decision shall be final and binding on the parties.).",
            "initial": False,
        },
        {
            "type": 3,
            "section": "2.1.18",
            "text": "**\"Property\"** means the Property as described in C of the Information Schedule.",
            "initial": False,
        },
        {
            "type": 3,
            "section": "2.1.19",
            "text": "**\"Purchaser\"** means the party/ies described as such in B of the Information Schedule.",
            "initial": False,
        },
        {
            "type": 3,
            "section": "2.1.20",
            "text": "**\"Rules\"** means the Management and Conduct Rules as Amended for the **\"{data["
                    "'development'].upper()}\"**  Sectional Title Scheme as prescribed in terms of Section 10(2)(a) "
                    "and (b) of the Sectional Titles Schemes Management Act No. 8 of 2011, subject to the approval "
                    "the Chief Ombud, and shall include any substituting rules, available on request from the Agent "
                    "or Seller.",
            "initial": False,
        },
        {
            "type": 3,
            "section": "2.1.21",
            "text": "**\"Scheme\"** means the **\"{data['development'].upper()}\"** Sectional Title Development to be "
                    "established on the Land, comprising of sectional title residential Units and Exclusive Use "
                    "Areas, which development may take place in phases and which is situated on the Land as depicted "
                    "on the Locality, Development, Unit and Exclusive Use Area Plans, Annexure's **\"A,"
                    "B & C\"** hereto",
            "initial": False,
        },
        {
            "type": 3,
            "section": "2.1.22",
            "text": "**\"Sectional Titles Act\"** means the Sectional Titles Act No 95 of 1986 (or any statutory "
                    "modification or re-enactment thereof) and includes the regulations made thereunder from time to "
                    "time.",
            "initial": False,
        },
        {
            "type": 3,
            "section": "2.1.23",
            "text": "**\"Sectional Plan\"** means the sectional plan/s prepared and registered in respect of the "
                    "Scheme and includes extension plans to be registered in respect of the Scheme.",
            "initial": False,
        },
        {
            "type": 3,
            "section": "2.1.24",
            "text": "**\"Sectional Titles Schemes Management Act\"** means the Sectional Titles Schemes Management "
                    "Act No. 8 of 2011 (or any statutory modification or re-enactment thereof) and includes the "
                    "regulations made thereunder from time to time.",
            "initial": False,
        },
        {
            "type": 3,
            "section": "2.1.25",
            "text": f"**\"Seller\"** means the party described in the Information Schedule. \n**\"Seller's "
                    f"attorneys\"** means Ilismi du Toit of LAÄS & SCHOLTZ INC, Queen Street Chambers, "
                    f"33 Queen Street, Durbanville, 7550, Tel (021) 975 0802 \nEmail: ilismi@lslaw.co.za, "
                    f"LAÄS & SCHOLTZ INC Attorneys Trust Bank Account details;  \nBank:  Standard Bank; Account "
                    f"Number: 272255505;  Branch Code: 051001; (Ref: Unit Number / {data['development'].upper()})",
            "initial": False,
        },
        {
            "type": 3,
            "section": "2.1.26",
            "text": "**\"Signature Date\"** means the date upon which this Agreement is signed by the party who signs "
                    "same last in time.",
            "initial": False,
        },
        {
            "type": 3,
            "section": "2.1.27",
            "text": "**\"Snags\"** means aesthetic and detail finishing items not affecting the Beneficial Occupation "
                    "of the Property.",
            "initial": False,
        },
        {
            "type": 3,
            "section": "2.1.28",
            "text": "**\"Defects List\"** means a list furnished by the Seller to the Purchaser on Occupation Date, "
                    "which list is to be completed by the Purchaser within 10 days after the Occupation Date, "
                    "where the Purchaser may identify construction items inside the Property that are to be attended "
                    "to by the Seller.",
            "initial": False,
        },
        {
            "type": 3,
            "section": "2.1.29",
            "text": "***\"Transfer Date\"** means the date of registration of transfer of the Unit in the name of the "
                    "Purchaser in the deeds office.",
            "initial": False,
        },
        {
            "type": 3,
            "section": "2.1.30",
            "text": "**\"Unit\"** means the residential Sectional Title Unit to be constructed by the Seller on the "
                    "Land for and on behalf of the Purchaser as envisaged herein.",
            "initial": False,
        },
        {
            "type": 3,
            "section": "2.1.31",
            "text": "**\"VAT\"** means value-added tax at the applicable rate in terms of the Value Added Tax Act No "
                    "89 of 1991 or any statutory re-enactment or amendment thereof.",
            "initial": False,
        },
        {
            "type": 3,
            "section": "2.1.32",
            "text": "**\"Works\"** means all the activities which are required to be undertaken to erect a "
                    "residential Unit on the Land for purposes of handover and transfer to the Purchaser.",
            "initial": False,
        },
        {
            "type": 2,
            "section": "2.2",
            "text": "The headnotes to the paragraphs in this Agreement are inserted for reference purposes only and "
                    "shall not affect the interpretation of any of the provisions to which they relate.",
            "initial": False,
        },
        {
            "type": 2,
            "section": "2.3",
            "text": f"Words importing the singular shall include the plural, and vice versa, and words importing the "
                    f"masculine gender shall include the feminine and neuter genders, and vice versa, "
                    f"and words importing persons shall include partnerships, trusts and bodies corporate, "
                    f"and *vice versa*.",
            "initial": False,
        },
        {
            "type": 2,
            "section": "2.4",
            "text": "If any provision in the Information Schedule, clause 1 and/or this clause 2 is a substantive "
                    "provision conferring rights or imposing obligations on any party, then notwithstanding that such "
                    "provision is contained in the Information Schedule, Clause 1 and/or this Clause 2  (as the case "
                    "may be) effect shall be given thereto as if such provision was a substantive provision in the "
                    "body of this Agreement.",
            "initial": False,
        },
        {
            "type": 1,
            "section": "3.",
            "text": "**SALE OF THE PROPERTY**",
            "initial": False,
        },
        {
            "type": 1,
            "section": "",
            "text": "The Seller hereby sells, and the Purchaser hereby purchases the Property, subject to and upon "
                    "the terms and conditions contained in this Agreement.",
            "initial": False,
        },
        {
            "type": 1,
            "section": "4.",
            "text": "**PURCHASE PRICE AND METHOD OF PAYMENT**",
            "initial": False,
        },
        {
            "type": 2,
            "section": "4.1",
            "text": "The Total purchase price of the Property shall be the amount stated in clause E1 of the "
                    "Information Schedule regardless of the final extent of the Unit as reflected on the Unit Plan "
                    "attached marked \"Annexure B\"",
            "initial": False,
        },
        {
            "type": 3,
            "section": "4.1.1",
            "text": "The Purchaser shall pay the Seller's attorneys the deposit for the Property as stated in clause "
                    "E2 of the Information Schedule within 3(three) days of signature of this Agreement by the "
                    "Purchaser, which deposit shall be held in trust by the Seller's attorneys and invested in an "
                    "interest-bearing account in accordance with the provisions of Section 26 of the Alienation of "
                    "Land Act No 68 of 1981 (as amended) with interest to accrue to the Purchaser.  The provisions of "
                    "this clause 4.2 shall constitute authority to the Seller's attorneys, in terms of Section 86(4) "
                    "of the Legal Practice Act, 2014(Act No. 28 of 2014), to invest the deposit for the benefit of "
                    "Purchaser pending registration of transfer.",
            "initial": False,
        },
        {
            "type": 3,
            "section": "4.1.2",
            "text": "Authority is hereby granted to the attorney by the Purchaser to withdraw the deposit as provided "
                    "for in clause 20.4 of the Agreement.",
            "initial": False,
        },
        {
            "type": 3,
            "section": "4.1.3",
            "text": "The Purchaser is aware that upon such withdrawal, no interest will be earned on the portion "
                    "being withdrawn.",
            "initial": False,
        },
        {
            "type": 2,
            "section": "4.2",
            "text": "**The Seller will not be bound to the Purchaser in terms of this Agreement until such time as "
                    "the deposit referred to in clause E2 has been paid to the Sellers attorneys trust account** "
                    "referred to in clause 4.1 above. The Seller shall be entitled to accept further offers "
                    "acceptable to the Seller, until such time as proof of payment of the deposit is furnished to the "
                    "Seller or the Seller's Attorneys, by the Purchaser, as provided for in this clause 4.2 In the "
                    "event of the Seller accepting an offer to purchase the Property on terms and conditions "
                    "acceptable to the Seller prior to receipt of such written notification, this Agreement shall be "
                    "deemed ipso facto null and void.",
            "initial": False,
        },
        {
            "type": 2,
            "section": "4.3",
            "text": "Within **21 (twenty one)** days after signature of this Agreement, the Purchaser shall furnish "
                    "the Seller or the Seller's Attorneys, with an irrevocable guarantee issued by a registered "
                    "commercial bank for the due payment of the balance of the purchase price of the Property as "
                    "referred to in clause E3 of the Information Schedule, or in the event of the Purchaser requiring "
                    "a mortgage bond for purposes of purchasing the Property in the amount recorded in clause E3 of "
                    "the Information Schedule, within **21 (twenty one)** days of securing a mortgage bond as "
                    "provided for in clause 19.1 hereunder. Should the Purchaser fail to comply with this clause 4.4, "
                    "the contract will be deemed null and void.",
            "initial": False,
        },
        {
            "type": 2,
            "section": "4.4",
            "text": "Or alternatively to the delivery of the guarantee referred to in clause 4.4 above, the Purchaser "
                    "shall within the same time periods as provided for in the aforesaid clause, pay into the trust "
                    "account of the Seller's attorneys, the balance of the purchase price of the Property as referred "
                    "to in clause E3 of the Information Schedule, to be held by such attorneys in an interest bearing "
                    "trust account, interest to accrue for the benefit of the Purchaser until the date upon which "
                    "payment of the relevant amount falls due to the Seller.  The Purchaser hereby irrevocably "
                    "authorises the attorneys to release from the funds so received, the payments due to the Seller "
                    "in terms of the provisions of this Agreement.",
            "initial": False,
        },
        {
            "type": 2,
            "section": "4.5",
            "text": "The Seller, at its sole discretion may elect to extend the periods as mentioned in clause 4.4 "
                    "and/or 4.5 above.",
            "initial": False,
        },
        {
            "type": 2,
            "section": "4.6",
            "text": "All amounts payable by the Purchaser in terms of this Agreement shall be paid to the Seller's "
                    "attorneys free of bank charges or commission at Cape Town and without deduction or set off by "
                    "means of a bank guaranteed cheque or a cheque drawn by a registered South African commercial "
                    "bank.",
            "initial": False,
        },
        {
            "type": 2,
            "section": "4.7",
            "text": "The Total Purchase Price of the Property as recorded in clause **E1** of the Information "
                    "Schedule shall be paid to the Seller on Transfer Date.",
            "initial": False,
        },
        {
            "type": 1,
            "section": "5.",
            "text": "**TRANSFER OF THE PROPERTY**",
            "initial": False,

        },
        {
            "type": 2,
            "section": "5.1",
            "text": "Transfer of the Property shall be passed by the Seller's attorneys and shall be given and taken "
                    "as close as possible to the Estimated Completion Date, recorded in clause D of the Information "
                    "Schedule.",
            "initial": False,

        },
        {
            "type": 2,
            "section": "5.2",
            "text": "TThe Seller shall be responsible for payment of the transfer costs of the Seller's attorneys "
                    "insofar as it relates to the transfer of the Property (plus VAT on such costs), costs of all "
                    "necessary affidavits and all other costs which have to be incurred in order to comply with the "
                    "statutes or other enactments or regulations relating to the passing of transfer of the Property.",
            "initial": False,
        },
        {
            "type": 2,
            "section": "5.3",
            "text": "The Seller shall be responsible for payment of all bond costs (if any) (plus VAT on such costs), "
                    "costs of all necessary affidavits initiation and valuation fees charged by the bank and all "
                    "other costs which have to be incurred in order to comply with the statutes or other enactments "
                    "or regulations relating to the mortgage of the Property.  However, the Purchaser will be liable "
                    "for the payment of initiation and/or valuation (bank administration) fees as may be charged by "
                    "the Bank.",
            "initial": False,
        },
        {
            "type": 2,
            "section": "5.4",
            "text": "The Purchaser shall within 7 (seven) days of being called upon to do so by the Sellers "
                    "attorneys, furnish all such information, sign all such documentation as may be necessary or "
                    "required to enable the Sellers attorneys to pass transfer of the Property and to register any "
                    "bond over the Property.",
            "initial": False,

        },
        {
            "type": 2,
            "section": "5.5",
            "text": "In particular the Purchaser must ensure that his tax affairs and the tax affairs of his "
                    "representatives, if applicable, are up to date as required by SARS to facilitate prompt issue by "
                    "SARS of the Transfer Duty Exemption.",
            "initial": False,
        },
        {
            "type": 2,
            "section": "5.6",
            "text": "Subject to the provisions of this Agreement, the Purchaser shall not, by reason of any delay in "
                    "the transfer of the Property to him due to any cause whatsoever, be entitled to cancel this "
                    "contract or to refrain from paying, or suspend payment of, any amount payable by him in terms of "
                    "this Agreement or to claim and recover from the Seller any damages or compensation or any "
                    "remission of rental.",
            "initial": False,
        },
        {
            "type": 1,
            "section": "6.",
            "text": "**RIGHTS OF SELLER:**",
            "initial": False,
        },
        {
            "type": 2,
            "section": "6.1",
            "text": "Pending establishment of the Body Corporate, the Seller shall be entitled to:",
            "initial": False,
        },
        {
            "type": 3,
            "section": "6.1.1",
            "text": "make Management and Conduct Rules for the use and enjoyment of the Common Property;",
            "initial": False,
        },
        {
            "type": 3,
            "section": "6.1.2",
            "text": "enter the Property at all reasonable times or to authorize it agents or workmen so to enter, "
                    "to inspect same or to carry out repairs;",
            "initial": False,
        },
        {
            "type": 3,
            "section": "6.1.3",
            "text": "o exercise all the rights and powers which a Body Corporate would be entitled to exercise in "
                    "terms of the Sectional Titles Act in respect of the Building, the Land and the owners and/or "
                    "occupants of Units.",
            "initial": False,
        },
        {
            "type": 3,
            "section": "6.1.4",
            "text": "cede rights for the generation of electricity created on the Property and or Common Property of "
                    "the Scheme to be registered on the Property, to a third party.",
            "initial": False,
        },
        {
            "type": 2,
            "section": "6.2",
            "text": "The Purchaser hereby appoints the Seller's nominee, irrevocably and in rem suam and with power "
                    "of substitution, to be his lawful agent and attorney:",
            "initial": False,
        },
        {
            "type": 3,
            "section": "6.2.1",
            "text": "to convene a meeting of the Body Corporate and there to vote in favour of any resolution of the "
                    "Body Corporate to amend the Rules or pass any other resolution as may be required:",
            "initial": False,
        },
        {
            "type": 3,
            "section": "6.2.2",
            "text": "by any bondholder for the grant of its consent to the opening of the sectional title register;",
            "initial": False,
        },
        {
            "type": 3,
            "section": "6.2.3",
            "text": "by the local or provincial authority and/or by a mortgagee prior to the grant of a sectional "
                    "mortgage bond over a Unit in the Scheme;",
            "initial": False,
        },
        {
            "type": 3,
            "section": "6.2.4",
            "text": "by the Seller in order to ensure the proper and efficient management and control of the scheme, "
                    "or to ensure that the Developer is able to exercise in full his rights to further develop the "
                    "Scheme.",
            "initial": False,
        },
        {
            "type": 3,
            "section": "6.2.5",
            "text": "to bind the Body Corporate to the pre-negotiated Services Contract concluded with a Managing "
                    "Agent.",
            "initial": False,
        },
        {
            "type": 3,
            "section": "6.2.6",
            "text": "to sign all documents necessary or required to comply with the Purchaser's obligations in terms "
                    "of this Agreement.",
            "initial": False,
        },
        {
            "type": 2,
            "section": "6.3",
            "text": "Conditions have been/will be imposed by the Seller (as Developer) in terms of Section 11(2) of "
                    "Act 95 of 1986 which conditions are/will be filed with the records of the Deeds Registry.",
            "initial": False,
        },
        {
            "type": 1,
            "section": "7.",
            "text": "**MORA INTEREST**",
            "initial": False,
        },
        {
            "type": 2,
            "section": "7.1",
            "text": "This Clause 7 shall be read in conjunction with clause 4.2 relating to PURCHASE PRICE AND METHOD "
                    "OF PAYMENT above, and clause 34 relating to **MORA INTEREST BREACHES BY Purchaser**.",
            "initial": False,
        },
        {
            "type": 2,
            "section": "7.2",
            "text": "In the event of there being any delay in connection with the registration of transfer for which "
                    "the Purchaser is responsible, then, without prejudice to any other rights of remedies the Seller "
                    "has in terms of this Agreement;",
            "initial": False,
        },
       {
            "type": 3,
            "section": "7.2.1",
            "text": "the Purchaser agrees to pay interest on the full purchase price at the prime interest rate plus "
                    "4% as certified by any commercial bank, from time to time calculated from the date the Purchaser "
                    "is notified in writing by the Seller (or the Seller's agent) as being in mora, to the date upon "
                    "which the Purchaser has ceased to be in mora, both days inclusive. A certificate by any Branch "
                    "Manager of such commercial bank, shall be *prima facie* proof of such prime interest rate.",
            "initial": False,
        },
        {
            "type": 3,
            "section": "7.2.2",
            "text": "The Purchaser shall be liable for a *pro rata* share of rates, taxes and other proprietary "
                    "charges payable in respect of the Property with effect from the Transfer Date, or Occupation "
                    "Date, whichever occurs first.",
            "initial": False,
        },
        {
            "type": 1,
            "section": "8.",
            "text": "**POSSESSION, OCCUPATION, RISK AND PROPRIETARY AND MUNICIPAL CHARGES**",
            "initial": False,
        },
        {
            "type": 2,
            "section": "8.1",
            "text": "The Seller shall give the Purchaser possession and occupancy of the Unit on the Transfer Date.",
            "initial": False,
        },
        {
            "type": 2,
            "section": "8.2",
            "text": "All risk and benefit in the Property shall be passed to the Purchaser on the Transfer Date.",
            "initial": False,
        },
        {
            "type": 2,
            "section": "8.3",
            "text": "The Purchaser shall be liable for a pro rata share of rates, taxes, levies and other proprietary "
                    "charges payable in respect of the Unit and Exclusive Use Areas with effect from the Transfer "
                    "Date, or Occupation Date, whichever occurs first.",
            "initial": False,
        },
        {
            "type": 2,
            "section": "8.4",
            "text": "The Purchaser shall refund to the Seller a proportionate share of such charges paid by the "
                    "Seller in advance and the Purchaser shall on demand pay to the Seller's attorneys an estimated "
                    "pro rata portion of such rates etc. in advance to enable the said Seller's attorneys to pay such "
                    "rates etc. before the Transfer Date.",
            "initial": False,
        },
        {
            "type": 2,
            "section": "8.5",
            "text": "Any deposits or payments paid by the Seller for and on behalf of the Purchaser shall be "
                    "refundable by the Purchaser immediately after the Seller has affected payment thereof.",
            "initial": False,
        },
        {
            "type": 2,
            "section": "8.6",
            "text": "In the event that occupation is given to the Purchaser, whether taken or not, prior to the "
                    "Transfer Date, the Purchaser shall be liable towards the Seller for occupational rental of **R12 "
                    "000.00 (twelve thousand rand)**, per month, payable from the Occupation Date (pro rata) in "
                    "advance towards the Seller. The Purchaser shall however not be entitled to take occupation of "
                    "the Property until such time as:-",
            "initial": True,
        },
        {
            "type": 3,
            "section": "8.6.1",
            "text": "the Sellers attorneys secured the full purchase price in cash or, in case of a mortgage loan, "
                    "a bank guarantee;",
            "initial": False,
        },
        {
            "type": 3,
            "section": "8.6.2",
            "text": "the Purchaser has signed all the Sellers attorneys documentation, including all transfer and "
                    "bond registration documentation;",
            "initial": False,
        },
        {
            "type": 3,
            "section": "8.6.3",
            "text": "the Purchaser (or his Agent) has signed the Happy Letter Document that the Unit is ready for "
                    "Beneficial Occupation.",
            "initial": False,
        },
        {
            "type": 2,
            "section": "",
            "text": "Once the aforesaid have been complied with, the Seller and/or his Agents will make arrangements "
                    "with the Purchaser for the delivery of the keys to the Unit.",
            "initial": False,
        },
        {
            "type": 2,
            "section": "8.7",
            "text": "The Seller shall notify the Purchaser once the City of Cape Town issued the Occupation "
                    "Certificate, certifying that the Unit is ready for Beneficial Occupation, **where after the "
                    "Purchaser shall, within 3 days of such notice, be obliged to inspect the Unit with the Seller or "
                    "his Agent, and sign the Happy Letter after such inspection**.",
            "initial": True,
        },
        {
            "type": 2,
            "section": "8.8",
            "text": "**In the event that the Purchaser delays the signing of the Happy Letter** by not attending to "
                    "the inspection with the Seller or his Agent **within the 3-day period** as per clause 8.7 above, "
                    "**then ipso facto, i.e. automatically, on the fourth calendar day the Purchaser will be deemed "
                    "to be in mora and liable to payment of interest on the full purchase price, as set out in clause "
                    "7.2.1 above**.",
            "initial": False,
        },
        {
            "type": 1,
            "section": "9.",
            "text": "**DEVELOPER: LIABILITY FOR DEFECTS**",
            "initial": False,
        },
        {
            "type": 2,
            "section": "9.1",
            "text": "The Property is sold subject to the conditions, reservations and servitudes contained in the "
                    "sectional title register and such conditions of sectional title as may be imposed by the "
                    "Developer, the City of Cape Town or any other authority.  Save as provided for in the **Consumer "
                    "Protection Act** and this Agreement to the contrary, the Purchaser purchases the Property "
                    "*\"voetstoots\"* and shall have no claim against the Seller respect of defects whether latent or "
                    "patent in the Property or the Common Property of the Scheme.",
            "initial": False,
        },
        {
            "type": 2,
            "section": "9.2",
            "text": "The Seller shall notify the Purchaser of the issue of the Occupation Certificate by the City of "
                    "Cape Town. The Seller (or his Agent) and Purchaser shall inspect the Unit and shall complete and "
                    "sign Happy Letter after such inspection.",
            "initial": False,
        },
        {
            "type": 2,
            "section": "9.3",
            "text": "The Purchaser shall in writing, within 10 days of taking occupation of the Unit, identify all "
                    "Defects in the Unit on the Defects List provided by Seller, where after the Seller shall attend "
                    "to such defects within 3 (three) months of the Purchaser taking occupation.",
            "initial": False,
        },
        {
            "type": 2,
            "section": "9.4",
            "text": "If there is any dispute regarding the existence or extent of any defect referred to in clause "
                    "9.3 above, the matter shall be referred to the Architect, whose decision shall be final and "
                    "binding upon the parties.",
            "initial": False,
        },
        {
            "type": 2,
            "section": "9.5",
            "text": "The Seller shall within a reasonable time remedy any defect in respect of roof leaks and gutter "
                    "leaks in the Building which may manifest themselves within 1 year after the Occupation Date "
                    "provided that the Purchaser notifies the Seller in writing within the said period of 1 year of "
                    "any such defects, failing which, the Purchaser shall be deemed to have accepted the Property in "
                    "such condition as at the Occupation Date.",
            "initial": False,
        },
        {
            "type": 2,
            "section": "9.6",
            "text": "The Seller shall within a reasonable time remedy any material structural defects in the Building "
                    "which may manifest themselves within 5 (five) years after the Occupation Date provided that the "
                    "Purchaser notifies the Seller in writing within the said period of 5 years of any such defects, "
                    "failing which, the Purchaser shall be deemed to have accepted the Property in such condition as "
                    "at the Occupation Date.",
            "initial": False,
        },
        {
            "type": 2,
            "section": "9.7",
            "text": "It is inevitable that hairline cracks occur. These cracks are not of a structural nature and are "
                    "caused by shrinkage and movement between building materials or settlement. These Cracks are "
                    "considered normal maintenance. The Seller shall not repair any cracking of this nature.",
            "initial": False,
        },
        {
            "type": 2,
            "section": "9.8",
            "text": "The Seller shall not be held responsible for any mould growth caused by lack of ventilation and "
                    "/ or condensation.",
            "initial": False,
        },
        {
            "type": 2,
            "section": "9.9",
            "text": "Approved windows and doors are used in construction, the Seller shall not be held responsible for doors and windows slamming in windy conditions or any damage caused thereby.",
            "initial": False,
        },
        {
            "type": 2,
            "section": "9.10",
            "text": "No window and door is completely weatherproof during heavy wind and storm conditions. Windows and doors are manufactured to comply with certain allowable tolerances and movement that will allow these to function in a proper manner under normal conditions. The Seller is not in control of this, and therefore shall not be held responsible for wind and rain that may enter under these conditions.",
            "initial": False,
        },
        {
            "type": 2,
            "section": "9.11",
            "text": "All warranties and undertakings given by the Seller to the Purchaser in terms of this Agreement are personal to the Purchaser who shall not be entitled to cede, assign or make over its rights thereto",
            "initial": False,
        },
        {
            "type": 1,
            "section": "10.",
            "text": "**BUILDING ON THE PROPERTY**",
            "initial": False,
        },
        {
            "type": 2,
            "section": "10.1",
            "text": "The Seller undertakes in a proper and workmanlike manner to erect the Unit and Exclusive Use Areas on the Land substantially in accordance with the Unit Plans and Specifications of Finishes attached hereto as Annexures **A to G**.",
            "initial": False,
        },
        {
            "type": 2,
            "section": "10.2",
            "text": "The Seller will supply all material and labour required for the Scheme. In the event of any discrepancy arising between the Plans and Specifications of Finishes, the provisions of the Specifications of Finishes shall prevail.",
            "initial": False,
        },
        {
            "type": 2,
            "section": "10.3",
            "text": "The Purchaser acknowledges that the Seller shall be entitled to appoint sub-contractors in respect of the whole or any part of the Scheme but shall notwithstanding such appointment remain liable to the Purchaser in terms hereof.",
            "initial": False,
        },
        {
            "type": 1,
            "section": "11.",
            "text": "**CONSTRUCTION OF THE PROPERTY AND THE SCHEME**",
            "initial": False,
        },
        {
            "type": 1,
            "section": "",
            "text": "The Seller shall erect a Unit and Scheme on the Land on the terms and conditions as provided for herein to include:-",
            "initial": False,
        },
        {
            "type": 2,
            "section": "11.1",
            "text": "The Unit and Scheme shall be erected and completed on the Land substantially in accordance with the Plans and Specifications annexed hereto as Annexure's **A to G** and initialled by both parties for identification purposes (\"the Unit and Scheme\").",
            "initial": False,
        },
        {
            "type": 2,
            "section": "11.2",
            "text": "Extras, variations, and omissions shall mean all work, which cannot reasonably be inferred from the Unit Plan and Specifications annexed hereto as Annexure1s **B and G**.",
            "initial": False,
        },
        {
            "type": 2,
            "section": "11.3",
            "text": "The Seller shall commence the building work within a reasonable period and shall complete the Unit as close as possible to the Estimated Occupation/Completion Date referred to in clause **D** of the Information Schedule and in accordance with the approved building plans and the Specifications.",
            "initial": False,
        },
        {
            "type": 2,
            "section": "11.4",
            "text": "The Purchaser will be given the opportunity to choose certain finishes from the selection to be made available by the Seller.  The Purchaser undertakes to finalise this choice of finishes within the time period granted by the Seller.  The Seller shall not guarantee the colour, texture, or availability of such finishes, and upon being advised by the Seller that the finishes selected are not available, the Purchaser shall forthwith choose alternatives thereto. **In the event that the Purchaser is unavailable to make the necessary choices, then the Seller will be authorised to make such choices in cases where delays are likely to be caused by the unavailability of the Purchaser**.",
            "initial": True,
        },
        {
            "type": 2,
            "section": "11.5",
            "text": "In the event of the building work being delayed by non-availability for any materials, plant or labour, accident on work site for which the Seller is not responsible, bad weather, viz major or other reasonable cause the Seller shall not be liable to the Purchaser for any damages caused by the delay.",
            "initial": True,
        },
        {
            "type": 2,
            "section": "11.6",
            "text": "If commencement of construction of the Unit is delayed for longer than a period of 6 months after the scheduled Commencement of Construction Date for any reason other than a reason attributable to the fault and/or omission of the Seller, then **the Seller shall be entitled in its sole discretion to resile from this Agreement**, with neither party having any further claim against one another, other than a refund to the Purchaser of the deposit paid in terms of clause **E2** of the Information Schedule, or alternatively claim an adjustment to the purchase price in accordance with any increases in the cost of material and/or labour which might in the interim have occurred.  In the event of the parties not being able to reach agreement as to the adjustment to the purchase price, then a quantity surveyor appointed by the Seller shall determine the dispute and the quantity surveyor's determination shall be final and binding on the parties",
            "initial": False,
        },
        {
            "type": 2,
            "section": "11.7",
            "text": "The Seller shall endeavour to complete the Unit by the Estimated Occupation Date referred to in clause **D** of the Information Schedule, subject to the provisions of clause 11.13 below. ",
            "initial": False,
        },
        {
            "type": 2,
            "section": "11.8",
            "text": "The Seller shall give to the Purchaser not less than 20 days' notice in writing of the anticipated Completion Date of the Unit, provided, however, that the Seller shall, after having given the Purchaser notice as contemplated aforesaid, be entitled to postpone the Completion Date by giving further notice to the Purchaser to this effect within 10 days after dispatch by the Seller of the first notice mentioned herein",
            "initial": False,
        },
        {
            "type": 2,
            "section": "11.9",
            "text": "For purposes aforesaid, the issue of an Occupancy Certificate by the building inspector employed by the relevant local authority or Architect shall constitute the Completion Date of the Unit or the day the Purchaser takes hand over of the keys, whichever is earlier as per clause 2.1.6 above.",
            "initial": False,
        },
        {
            "type": 2,
            "section": "11.10",
            "text": "The occurrence of the event detailed in clause 11.9 above shall constitute complete proof of the satisfactory completion of the Unit by the Seller and the Seller shall be discharged completely from all obligations expressed or implied under this Agreement and any variation thereof or addition thereto and the Purchaser shall have no further claim on the Seller, save as specifically otherwise provided herein.",
            "initial": False,
        },
        {
            "type": 2,
            "section": "11.11",
            "text": "Notwithstanding anything elsewhere provided for in this Agreement all amounts owing in terms of this Agreement which have not already been paid in terms of the provisions of this Agreement shall be forthwith payable on the Completion Date.",
            "initial": False,
        },
        {
            "type": 2,
            "section": "11.12",
            "text": "The Purchaser acknowledges that any extension of period granted by the Seller for the fulfilment of any suspensive conditions in term of this Agreement may cause a delay of the Estimated Completion Date. The Seller shall not be held liable for any delays caused by such extensions.",
            "initial": False,
        },
        {
            "type": 2,
            "section": "11.13",
            "text": "The Estimated Completion Date is 4 to 5 months after the commencement of the construction date of this Agreement, **provided that any and/or all suspensive conditions of this Agreement has been fulfilled**, excluding any delays provided for in terms of this Agreement.",
            "initial": False,
        },
        {
            "type": 2,
            "section": "11.14",
            "text": "The Purchaser shall not be allowed on the Land, nor permitted to interfere with building operations on the Land or issue any instructions to the Seller's contractor/s or subcontractor/s during the construction period, without specific written authorization from the Seller.",
            "initial": False,
        },
        {
            "type": 1,
            "section": "12.",
            "text": "**SITING AND FINISHES OF BUILDINGS**",
            "initial": False,
        },
        {
            "type": 2,
            "section": "12.1",
            "text": "Should it for any reason in the Seller's sole discretion be required, **the Seller shall be entitled to make the necessary changes** where it in its sole discretion considers it necessary to **alter the siting of the Unit, Exclusive Use Areas or out-buildings** from the positions shown on the Development Plan and drawings forming part of the Annexure's hereto, subject to the condition that any additional costs incurred in making these alterations shall be borne by the Seller.",
            "initial": False,
        },
        {
            "type": 2,
            "section": "12.2",
            "text": "In the event of an error in the siting of the Unit, Exclusive Use Areas or out-buildings by the Seller such error shall not be deemed to constitute a breach of this Agreement by the Seller and the Seller shall have the right and the Purchaser hereby automatically authorises the Seller to make such amendments, alterations or modifications to the Plans and/or Specifications and/or the Unit, Exclusive Use Areas or outbuildings as may be necessary in order to legitimise the erroneous siting thereof or if necessary to re-site the same so as to comply with any law, bylaw, regulation, condition of title or the like, which would otherwise have been breached by such erroneous siting's of the Unit, Exclusive Use Areas or out-buildings.",
            "initial": False,
        },
        {
            "type": 2,
            "section": "12.3",
            "text": "The Purchaser shall have no claim of whatsoever nature or howsoever arising against the Seller for damages as a result of a change of or an error in the siting of the Unit, Exclusive Use Areas or out-buildings",
            "initial": False,
        },
        {
            "type": 2,
            "section": "12.4",
            "text": "The placement of any boundary wall to the Land, is determined by the Architect, accordingly such boundary wall may, or may not, form part of the Land.",
            "initial": False,
        },
        {
            "type": 1,
            "section": "13.",
            "text": "**VARIATIONS**",
            "initial": False,
        },
        {
            "type": 2,
            "section": "13.1",
            "text": "The Seller shall not be obliged to agree to any additional works or changes to the Unit as may be requested by the Purchaser.  Any agreed changes as may be required by the Purchaser must be paid by the Purchaser to the Seller on request, otherwise the works shall proceed as per the Specifications of Finishes and Plans.  The Seller shall not be liable for any delay in the Completion Date of the Unit should such delay be attributed to variations required by the Purchaser.  The Purchaser shall further be liable for any and all costs attributed to delays caused by the Purchaser, which delays, and costs will be communicated to the Purchaser, by the Seller, prior to commencement of such works.",
            "initial": False,
        },
        {
            "type": 2,
            "section": "13.2",
            "text": "The Purchaser shall not under any circumstances be permitted to issue any instructions directly to the building contractors.  All matters related to the Unit and Scheme shall be directed to the appointed Agent of the Seller in writing.",
            "initial": False,
        },
        {
            "type": 1,
            "section": "14.",
            "text": "**RULES AND TITLE CONDITIONS**",
            "initial": False,
        },
        {
            "type": 2,
            "section": "14.1",
            "text": "The Property is sold subject to such Rules, conditions, restrictions, servitudes and other provisions set out or referred to in the Title Deed and / or approved sectional title plan of the aforementioned Scheme, and all such conditions, servitudes and /or restrictions and/or changes that may be imposed by the Seller, any professional consultant of the Seller or any local or regional authority as a condition of rezoning, subdivision or building plan approval. **The Purchaser/s hereby confirms that he/she will familiarize him/herself with the Rules, to be approved by the Ombud, and should the Purchaser not agree/accept the said Rules, the Purchaser shall address such concerns with the trustees of the Body Corporate or the Ombud once the Purchaser is a member of the Scheme**.",
            "initial": True,
        },
        {
            "type": 2,
            "section": "14.2",
            "text": "The Purchaser acknowledges that he has acquainted himself with the nature, condition, and locality of the Property and Scheme. The Purchaser will have no claim whatsoever against the Seller for any deficiency in the size of the Property within a variance of 5%, which may be revealed on any re-survey nor shall the Seller benefit from any possible excess.",
            "initial": False,
        },
        {
            "type": 2,
            "section": "14.3",
            "text": "Should there be deficiency in the size of the Property, exceeding a variance of 5% referred to in Clause 14.2 above, the Seller shall be entitled to amend the Purchase Price, by the cost per square metre, as calculated in this Agreement, for the square meterage of the final extent of the Property, which is below the 5% variance.",
            "initial": False,
        },
        {
            "type": 2,
            "section": "14.4",
            "text": "Notwithstanding anything previously provided, the Seller shall under no circumstances be responsible for damage and/or loss caused by wear and tear, misuse, neglect, negligence, abuse, accident or in respect of any matter arising from or relating to a risk insured against in terms of Homeowners Insurance Policies normally issued by a South African Insurance Company in respect of residential properties.  The Seller shall furthermore under no circumstances be liable for any consequential loss or damages.",
            "initial": False,
        },
        {
            "type": 2,
            "section": "14.5",
            "text": "In the event of there being any dispute between the parties as regards the matter of whether any item complained of by the Purchaser constitutes a defect covered by the guarantee and/or any dispute relating to the repair of the defect, such dispute will be determined by the Architect whose determination shall be final and binding on the parties",
            "initial": False,
        },
        {
            "type": 2,
            "section": "14.6",
            "text": "Such guarantees as may be received by the Seller in respect of any item incorporated in the Property shall, to the extent that the Seller is entitled to do so, be passed on to the Purchaser with effect from the Completion Date.",
            "initial": False,
        },
        {
            "type": 2,
            "section": "14.7",
            "text": "Upon transfer, the Purchaser shall automatically become a member of the GREY HERON HOME OWNER'S ASSOCIATION,  and a condition of the title to this effect, will be included in the title deed of the property.",
            "initial": False,
        },
        {
            "type": 1,
            "section": "15.",
            "text": "**VALUE-ADDED TAX**",
            "initial": False,
        },
        {
            "type": 2,
            "section": "15.1",
            "text": "Unless the context of the clause concerned clearly indicates that the amount concerned is exclusive of VAT, all amounts provided for in this Agreement shall be inclusive of VAT.",
            "initial": False,
        },
        {
            "type": 2,
            "section": "15.2",
            "text": "All or any VAT payable by the Purchaser in terms of this Agreement arising from the supply of any goods and/or services (as defined in the Value-Added Tax Act No 89 of 1991 or any statutory re-enactment or modification thereof) by the Seller to the Purchaser in terms of this Agreement shall become due for payment and shall be paid by the Purchaser forthwith upon presentation of the relevant invoice by the Seller to the Purchaser.",
            "initial": False,
        },
        {
            "type": 2,
            "section": "15.3",
            "text": "Any dispute which may arise between the Seller and the Purchaser as to the liability for and/or payment of VAT or the amount thereof in terms of Clause 15.2 above shall be referred to the auditors of the Seller for the time being for decision and their decision shall be final and binding as between the parties and carried into effect.",
            "initial": False,
        },
        {
            "type": 1,
            "section": "16.",
            "text": "**COMPLIANCE WITH STATUTES AND BY-LAWS**",
            "initial": False,
        },
        {
            "type": 1,
            "section": "",
            "text": "The Seller shall ensure that the Unit and Scheme is constructed in conformity with the provisions of any legislation in force affecting the said Unit and will give all necessary notices to, and obtain the requisite sanction of, the local authority, in respect of the said Unit and generally ensure that the building and other regulations of such authority be complied with.",
            "initial": False,
        },
        {
            "type": 1,
            "section": "17.",
            "text": "**PUBLIC LIABILITY INSURANCE**",
            "initial": False,
        },
        {
            "type": 1,
            "section": "",
            "text": "The Seller or alternatively, the appointed contractor to attend to the construction of the Unit and Scheme, shall reasonably insure against public liability on or around the Unit and Scheme from the commencement of building operations until completion of the Unit in terms of this Agreement and until the risk in the Unit has passed to the Purchaser in terms of this Agreement.",
            "initial": False,
        },
        {
            "type": 1,
            "section": "18.",
            "text": "**NATIONAL HOMEBUILDERS REGISTRATION COUNCIL**",
            "initial": False,
        },
        {
            "type": 2,
            "section": "18.1",
            "text": "It is recorded that the Seller is registered with the National Homebuilders Registration Council (NHBRC Registration Number: 3000179685) and that the Unit shall be enrolled by his appointed contractor CAPE PROJECTS CONSTRUCTION (Pty) (NHBRC Registration Number: 3000156495).",
            "initial": False,
        },
        {
            "type": 2,
            "section": "18.2",
            "text": "The registration levy to be payable to the National Homebuilders Registration Council arising from the aforementioned registration shall be paid by the Seller.",
            "initial": False,
        },
        {
            "type": 2,
            "section": "18.3",
            "text": "The building shall be constructed as per the guidelines prescribed in the Housing Consumer Protection Measures Act, 1998 (as amended), and National Building Regulations where applicable.",
            "initial": False,
        },
        {
            "type": 1,
            "section": "19.",
            "text": "**BOND**",
            "initial": False,
        },
        {
            "type": 1,
            "section": "",
            "text": "Should the Purchaser signify on the Information Schedule that he shall require a loan to part finance the acquisition of the Unit against the security of the mortgage bond to be registered over the Unit, then:",
            "initial": False,
        },
        {
            "type": 2,
            "section": "19.1",
            "text": "This Agreement is subject to the **Purchaser obtaining the approval of a loan in principle** from a bank or other recognized financial institution for the amount (if any) stated in clause **E3** of the Information Schedule within **21 (twenty one)** days of the Signature Date or such extended period as the Seller in its sole discretion may determine.",
            "initial": False,
        },
        {
            "type": 2,
            "section": "19.2",
            "text": "Should the Seller elect to extend the period within which its loan is to be granted, the Seller and/or its agent shall be entitled to apply for the loan to any financial institution on behalf of the Purchaser and the Purchaser hereby grants to the Seller and/or its agent an irrevocable power of attorney in *rem suam* to make application on its behalf in this regard for the duration of the extended period; ",
            "initial": False,
        },
        {
            "type": 2,
            "section": "19.3",
            "text": "**The Purchaser agrees to make use of the services of MORTGAGE MAX, Tel no. 021 913 1944 (Sophia Vorster Tel no.  082 372 8074) as the mortgage originator for the loan referred to**",
            "initial": True,
        },
        {
            "type": 2,
            "section": "19.4",
            "text": "**If the Purchaser intends on making use of a private banker, the Purchaser undertakes to inform the mortgage originator with the name and contact details of such private banker, within 3 days of Signature Date.  In such case, the bond costs will not be paid by the Seller**.",
            "initial": True,
        },
        {
            "type": 2,
            "section": "19.5",
            "text": "The Purchaser undertakes to sign all documents and do all things necessary to ensure the successful granting of the loan.  Without derogating from the generality of the aforegoing, the Purchaser shall make a written application for the loan within 3 days after Signature Date and should such application be unsuccessful, the Purchaser shall, until the expiry of the initial period or the extended period (as the case may be) nevertheless continue to use its best endeavours and to do all things that may be necessary in order to obtain the loan elsewhere;",
            "initial": False,
        },
        {
            "type": 2,
            "section": "19.6",
            "text": "**All costs to be associated with the registration of the mortgage bond to secure the loan to be taken up by the Purchaser shall be for the account of the Seller, unless otherwise herein stipulated. Furthermore, the Seller shall not pay for a bond costs in the instance of the Purchaser making use of a private banker and/or own originator.**",
            "initial": True,
        },
        {
            "type": 2,
            "section": "19.7",
            "text": "Upon the issue to the Purchaser by the said financial institution of a written quotation and a written pre  agreement statement (as contemplated in Section 92 of the National Credit Act, No 34 of 2005) in respect of the mortgage loan in the said amount recorded in **E3** of the Information Schedule **whether or not such quotation or pre-agreement statement is accepted by the Purchaser, the mortgage loan shall be deemed to have been approved**.",
            "initial": False,
        },
        {
            "type": 2,
            "section": "19.8",
            "text": "The Purchaser shall accept the pre  agreement statement (as contemplated in Section 92 of the National Credit Act, No 34 of 2005) in respect of the mortgage loan in the said amount recorded in **E3** of the Information Schedule within 7 days of issue to the Purchaser by the said financial institution.",
            "initial": True,
        },
        {
            "type": 2,
            "section": "19.9",
            "text": "**In the event that the Purchaser delays the signing of the Acceptance of the pre  agreement statement (as contemplated in Section 92 of the National Credit Act, No 34 of 2005) within the 7-day period** as per clause 19.8 above, **then *ipso facto**, i.e. automatically, on the eighth calendar day the Purchaser will be deemed to be in mora and liable to payment of interest on the full purchase price, as set out in clause 7.2.1 above.**",
            "initial": True,
        },
        {
            "type": 1,
            "section": "20.",
            "text": "**BROKERAGE**",
            "initial": False,
        },
        {
            "type": 2,
            "section": "20.1",
            "text": "The parties' record that the Agent named in the Information Schedule was the effective cause of this transaction.",
            "initial": False,
        },
        {
            "type": 2,
            "section": "20.2",
            "text": "The Seller shall pay the brokerage to the said Agent in accordance with the terms of the mandate granted to the Agent by the Seller.",
            "initial": False,
        },
        {
            "type": 2,
            "section": "20.3",
            "text": "The Purchaser warrants and undertakes to the Seller that neither the Seller nor the Unit was introduced to the Purchaser by any party other than the Agent referred to in Clause 20.1 above and indemnifies the Seller against any claim for commission arising from any breach of this warranty.",
            "initial": False,
        },
        {
            "type": 2,
            "section": "20.4",
            "text": "The parties agree that the Agent shall be entitled to a part payment of commission, equal to the deposit paid by the purchaser upon:",
            "initial": False,
        },
        {
            "type": 3,
            "section": "20.4.1",
            "text": "the payment of the balance purchase price or the delivery of a guarantee acceptable to the attorney; or",
            "initial": False,
        },
        {
            "type": 3,
            "section": "20.4.2",
            "text": "receipt by the attorney of a bond instruction for registration of a mortgage bond as per clause 19 above",
            "initial": False,
        },
        {
            "type": 2,
            "section": "20.5",
            "text": "In the event that this Agreement is cancelled by agreement, or by the Purchaser due to any default by the Seller, the Seller undertakes to immediately pay this portion of the deposit together with the balance held by the Attorney into a nominated account of the Purchaser, without deduction or set-off.",
            "initial": False,
        },
        {
            "type": 1,
            "section": "21.",
            "text": "**CANCELLATION**",
            "initial": False,
        },
        {
            "type": 2,
            "section": "21.1",
            "text": "In the event of the Purchaser refusing or failing to comply punctually with any of his obligation in terms of this Agreement and provided that 7 (seven) days has elapsed after receipt by the Purchaser of a written demand to comply with the said obligation(s), the Seller shall be entitled to:",
            "initial": False,
        },
        {
            "type": 3,
            "section": "21.1.1",
            "text": "cancel this Agreement, and to retain all monies paid as *\"rouwkoop\"* and liquidated damages, without prejudice to his rights to claim damages from the Purchaser; or",
            "initial": False,
        },
        {
            "type": 3,
            "section": "21.1.2",
            "text": "to claim specific performance from the Purchaser, i.e. that the Purchaser complies with all his obligations in terms of the Agreement, including, but not limited to payment of all legal costs, pro rata rates, taxes and levies.",
            "initial": False,
        },
        {
            "type": 2,
            "section": "21.2",
            "text": "In the event of the Purchaser being provisionally or finally sequestrated or liquidated, the Seller shall enjoy the same rights as set out above.",
            "initial": False,
        },
        {
            "type": 1,
            "section": "22.",
            "text": "**ARBITRATION**",
            "initial": False,
        },
        {
            "type": 2,
            "section": "22.1",
            "text": "Any dispute, question or difference arising at any time between the parties to this Agreement out of or in regard to any matters arising out of; or the rights and duties of any of the parties hereto; or the interpretation of; or the termination of; or any matter arising out of the termination of; or the rectification of this Agreement, shall be submitted to and decided by arbitration on notice given by either party to the other of them in terms of this clause.",
            "initial": False,
        },
        {
            "type": 2,
            "section": "22.2",
            "text": "There will be one arbitrator who will be, if the question in issue is:",
            "initial": False,
        },
        {
            "type": 3,
            "section": "22.2.1",
            "text": "primarily a legal matter, a practicing advocate or attorney of not less than ten years' standing;",
            "initial": False,
        },
        {
            "type": 3,
            "section": "22.2.2",
            "text": "primarily a technical matter, an architect or quantity surveyor, depending on the nature of the dispute.",
            "initial": False,
        },
        {
            "type": 2,
            "section": "22.3",
            "text": "The appointment of the arbitrator will be agreed upon between the parties to the dispute but failing agreement between them within a period of fourteen days after the arbitration has been demanded, any of the parties to the dispute shall be entitled to request the Chairman for the time being of the Cape Bar Council to make the appointment and who, in making his appointment, will have regard to the nature of the dispute.",
            "initial": False,
        },
        {
            "type": 2,
            "section": "22.4",
            "text": "The arbitrator shall have the powers conferred upon an arbitrator in the Arbitration Act No. 42 of 1965, as amended or re-enacted in some other form from time to time but will not be obliged to follow the procedures described in that Act and will be entitled to decide on such procedures as he may consider desirable for the speedy determination of the dispute.",
            "initial": False,
        },
        {
            "type": 2,
            "section": "22.5",
            "text": "The arbitration shall be held in Cape Town in accordance with the provision of the Arbitration Act referred to above, save that the arbitration shall be informal and the parties shall not be entitled to legal representation but shall be represented solely by themselves or in the case of a company or a business, by a member or members of their full-time management or of their boards of directors, it being the agreed intention that, if possible, the arbitration shall be held and concluded as soon as is reasonably practical after it has been demanded.",
            "initial": False,
        },
        {
            "type": 2,
            "section": "22.6",
            "text": "The decision of the arbitrator, including any order as to the costs of the arbitration, shall be final and binding on the parties and may be made an order of any court of competent jurisdiction.",
            "initial": False,
        },
        {
            "type": 2,
            "section": "22.7",
            "text": "This clause is severable from the rest of the Agreement and shall therefore remain in effect even if this Agreement is terminated.",
            "initial": False,
        },
        {
            "type": 1,
            "section": "23.",
            "text": "**CAPACITY OF PURCHASER**",
            "initial": False,
        },
        {
            "type": 1,
            "section": "",
            "text": "If the Purchaser signs this Agreement as a principal in terms of a contract for the benefit of a "
                    "third party, the latter being a company or close corporation to be incorporated the Purchaser "
                    "in his personal capacity shall be regarded as the Purchaser in terms of this Agreement unless "
                    "the said company or close corporation is incorporated and duly adopts and ratifies this Agreement "
                    "within 60 (sixty) days after the date upon which the Seller signs this Agreement. "
                    "In the event of the said company or close corporation being duly incorporated and adopting and "
                    "ratifying this Agreement in terms as set out above then the Purchaser, by his signature hereto, "
                    "hereby interposes and binds himself in favour of the Seller as surety and co-principal debtor "
                    "in solidum with such company or close corporation for the due and timeous performance by it of "
                    "all of its obligations as Purchaser in terms of this Agreement.",
            "initial": False,
        },
        {
            "type": 1,
            "section": "24.",
            "text": "**GENERAL**",
            "initial": False,
        },
        {
            "type": 2,
            "section": "24.1",
            "text": "This Agreement constitutes the sole and entire agreement between the parties and no warranties, representations, guarantees or other terms and conditions of whatsoever nature not contained or recorded herein shall be of any force or effect.",
            "initial": False,
        },
        {
            "type": 2,
            "section": "24.2",
            "text": "No variation of the terms and conditions of this Agreement or any consensual cancellation thereof shall be of any force or effect unless reduced to writing and agreed by the parties or their duly authorized representatives.",
            "initial": False,
        },
        {
            "type": 2,
            "section": "24.3",
            "text": "No indulgence, which the Seller may grant to the Purchaser, shall constitute a waiver of any of the rights of the Seller who shall not thereby be precluded from exercising any rights against the Purchaser which may have risen in the past or which might arise in the future.",
            "initial": False,
        },
        {
            "type": 2,
            "section": "24.4",
            "text": "The Purchaser warrants that all consents required in terms of the Matrimonial Property Act No. 88 of 1984 have been duly furnished.",
            "initial": False,
        },
        {
            "type": 2,
            "section": "24.5",
            "text": "It is hereby recorded that the headings to the clauses in this Agreement are inserted for information only and will have no relevance in the interpretation thereof.  The singular shall be deemed to include the plural (and vice versa) and the one sex the other.",
            "initial": False,
        },
        {
            "type": 1,
            "section": "25.",
            "text": "**PHASED DEVELOPMENT**",
            "initial": False,
        },
        {
            "type": 2,
            "section": "25.1",
            "text": "The Purchaser acknowledges that the Seller intends to extend the Scheme by erecting and completing from time-to-time further buildings on specified parts of the Common Property, to divide such Buildings into sections and Common Property and confer the right of exclusive use over parts of such Common Property upon the owner or owners of one or more of such sections and to reserve its right in this regard in accordance with provisions of section 25(1) of the Sectional Titles Act",
            "initial": False,
        },
        {
            "type": 2,
            "section": "25.2",
            "text": "The Purchaser shall be obliged to allow the Seller or its successor in title (\"the Developer\") to exercise its right to develop the sections in the manner envisaged herein and shall not be entitled to interfere with or obstruct the Developer in any way from erecting the said Buildings on the Common Property.",
            "initial": False,
        },
        {
            "type": 1,
            "section": "26.",
            "text": "**RE-SALE OF PROPERTY**",
            "initial": False,
        },
        {
            "type": 1,
            "section": "",
            "text": "In order to ensure the Purchaser does not compete with, nor impede the Seller in the marketing and sales of the Units, the Purchaser shall not be entitled to market and sell the Property prior to the Seller having sold the entire Scheme to end users.",
            "initial": False,
        },
        {
            "type": 1,
            "section": "27.",
            "text": "**CERTIFICATE OF COMPLIANCE OF WATER INSTALLATION AND ELECTRICAL CERTIFICATE**",
            "initial": False,
        },
        {
            "type": 1,
            "section": "",
            "text": "The Seller shall, before Transfer Date and at his expense, submit a Certificate from an accredited plumber to the City of Cape Town Municipality, certifying that the water supply to the Unit confirms with the requirements stipulated in Section 14 of the City of Cape Town: Water By-law, 2010, namely that:",
            "initial": False,
        },
        {
            "type": 2,
            "section": "(a)",
            "text": "the water installation conforms to the National Building Regulations and this By-law;",
            "initial": False,
        },
        {
            "type": 2,
            "section": "(b)",
            "text": "there are no defects which can cause water to run to waste;",
            "initial": False,
        },
        {
            "type": 2,
            "section": "(c)",
            "text": "the water meter registers; and",
            "initial": False,
        },
        {
            "type": 2,
            "section": "(d)",
            "text": "there is no discharge of storm water into the sewer system.",
            "initial": False,
        },
        {
            "type": 1,
            "section": "",
            "text": "The Seller undertakes to submit the said Certificate to the City of Cape Town via fax or email, and to furnish proof of such submission to the Sellers attorneys. Insofar as the accredited plumber appointed by the Seller to provide such Certificate requires corrective work to be carried out as a precondition to the issue thereof, the Seller will procure such work is carried out at his cost and expense. The Seller further undertakes to furnish the Purchaser, prior to transfer, with a certificate of compliance issued by an accredited person, declaring that the electrical installation in the Unit (up to and including the distribution board) complies with the provisions of Regulation 4(1) of the Electrical Installation Regulations of the Machinery and Occupational Safety Act.",
            "initial": False,
        },
        {
            "type": 1,
            "section": "28.",
            "text": "**JURISDICTIONN**",
            "initial": False,
        },
        {
            "type": 1,
            "section": "",
            "text": "The parties hereby consent in terms of Section 45 of the Magistrate's Court Act No. 32 of 1944, as amended, to the jurisdiction of the Magistrate's Court of any district having jurisdiction in terms of Section 28(1) of the said Magistrate's Court Act in any action or court procedure instituted by the Seller arising out of this Agreement. Notwithstanding the above, the Seller shall be entitled to institute any action or court procedure against the Purchaser arising out of this Agreement in any Court having jurisdiction.",
            "initial": False,
        },
        {
            "type": 1,
            "section": "29.",
            "text": "**72 HOUR CLAUSE**",
            "initial": False,
        },
        {
            "type": 1,
            "section": "",
            "text": "The Seller can, before the suspensive conditions in this Agreement (first transaction) are fulfilled, accept subsequent offer(s); which will not be subject to any suspensive conditions or whereof all suspensive conditions have already been fulfilled. The Purchaser in respect of the first transaction then has 72 hours in which to comply with the suspensive conditions in the first transaction. The 72 hours are not applicable during weekends and public holidays. The 72 hours commence when the Purchaser or his agent in respect of the first transaction:",
            "initial": False,
        },
        {
            "type": 2,
            "section": "1)",
            "text": "Is notified in writing of any subsequent offer(s) between 08h00 and 17h00;",
            "initial": False,
        },
        {
            "type": 2,
            "section": "2)",
            "text": "Receive a copy of the subsequent offer(s);",
            "initial": False,
        },
        {
            "type": 2,
            "section": "3)",
            "text": "Receive proof that all suspensive conditions of the subsequent offer(s) have been fulfilled, including bond approval with conditions acceptable for the SELLER; and",
            "initial": False,
        },
        {
            "type": 2,
            "section": "4)",
            "text": "Any shortfall on the Purchaser price of the subsequent offer(s), not covered by cash or bond (if applicable), is secured by a guarantee delivered to the Seller's attorneys.",
            "initial": False,
        },
        {
            "type": 1,
            "section": "30.",
            "text": "**CO-OPERATION**",
            "initial": False,
        },
        {
            "type": 2,
            "section": "30.1",
            "text": "Each of the parties hereby undertakes to:",
            "initial": False,
        },
        {
            "type": 3,
            "section": "30.1.1",
            "text": "sign and/or execute all such documents (and without limiting the generality of the aforegoing, same shall include the execution of the necessary power of attorney and transfer duty declarations);",
            "initial": False,
        },
        {
            "type": 3,
            "section": "30.1.2",
            "text": "do and to procure the doing by other persons, and to refrain and procure that other persons will refrain from doing, all such acts; and",
            "initial": False,
        },
        {
            "type": 3,
            "section": "30.1.3",
            "text": "pass, and to procure the passing of all such resolutions of directors or shareholders of any company, or members of any close corporation, or trustees of any trust, as the case may be;",
            "initial": False,
        },
        {
            "type": 2,
            "section": "",
            "text": "to the extent that the same may lie within the power of such party and may be required to give effect to the import or intent of this Agreement, and any contract concluded pursuant to the provisions of this Agreement.",
            "initial": False,
        },
        {
            "type": 2,
            "section": "30.2",
            "text": "The Purchaser undertakes to sign all necessary transfer and bond documentation and to pay all costs relating thereto within 7 days of the date of despatch of written notice from the Seller's attorneys to do so.",
            "initial": False,
        },
        {
            "type": 1,
            "section": "31.",
            "text": "**NOTICES AND DOMICILIA**",
            "initial": False,
        },
        {
            "type": 2,
            "section": "31.1",
            "text": "Each party chooses as his *domicilium citandi et executandi* his address as set out in the Information Schedule, at which address all notices and legal processes in relation to this Agreement or any action arising therefrom may be effectually delivered and served.",
            "initial": False,
        },
        {
            "type": 2,
            "section": "31.2",
            "text": "Any notice given by one of the parties to the other (the \"addressee\") which:",
            "initial": False,
        },
        {
            "type": 3,
            "section": "31.2.1",
            "text": "is delivered by hand to the addressee's *domicilium citandi et executandi* shall be presumed, until the contrary is proved, to have been received by the addressee on the date of delivery; or",
            "initial": False,
        },
        {
            "type": 3,
            "section": "31.2.2",
            "text": "is posted by prepaid registered post from an address within the Republic of South Africa to the addressee at the addressee's *domicilium citandi et executandi* shall be presumed, until the contrary is proved, to have been received by the addressee on the fifth day after the date of posting; or",
            "initial": False,
        },
        {
            "type": 3,
            "section": "31.2.3",
            "text": "is delivered by fax and / or email shall be presumed, until the contrary is proved, to have been received by the addressee on the date of delivery.",
            "initial": False,
        },
        {
            "type": 2,
            "section": "31.3",
            "text": "Either party shall be entitled, on written notice to the other, to change the address of his *domicilium citandi et executandi*.",
            "initial": False,
        },
        {
            "type": 1,
            "section": "32.",
            "text": "**DIRECT MARKETING AND COOLING OFF PERIOD**",
            "initial": False,
        },
        {
            "type": 2,
            "section": "32.1",
            "text": "In complying with the Consumer Protection Act, certain portions of the Agreement have been printed in bold. The reason for this is to specifically draw the Purchaser's attention to these paragraphs as they either: ",
            "initial": False,
        },
        {
            "type": 3,
            "section": "32.1.1",
            "text": "limit in some way the risk or liability of the Seller or any other person;",
            "initial": False,
        },
        {
            "type": 3,
            "section": "32.1.2",
            "text": "constitute an assumption of risk or liability by the Purchaser;",
            "initial": False,
        },
        {
            "type": 3,
            "section": "32.1.3",
            "text": "impose an obligation on the Purchaser to indemnify the Seller or any other person for some cause; or",
            "initial": False,
        },
        {
            "type": 3,
            "section": "32.1.4",
            "text": "is an acknowledgement of a fact by the Purchaser.",
            "initial": False,
        },
        {
            "type": 2,
            "section": "32.2",
            "text": "The Purchaser, in the event of having concluded this Agreement as a result of Direct Marketing as defined in the provisions of the Consumer Protection Act No. 68 of 2008, confirms that he/she/it has been informed of his rights as provided for in Section 16 read with Section 20 (2) (a) of the aforementioned Act, to rescind a transaction, without reason or penalty, within 5 (five) business days after the later of the date on which:",
            "initial": False,
        },
        {
            "type": 3,
            "section": "32.2.1",
            "text": "the transaction or Agreement was signed; or",
            "initial": False,
        },
        {
            "type": 3,
            "section": "32.2.2",
            "text": "the goods that were the subject of the transaction were delivered to the consumer.",
            "initial": False,
        },
        {
            "type": 2,
            "section": "32.3",
            "text": "Further to the provisions of clause 33.2 above, **the Purchaser hereby warrants that this Agreement has not been concluded as a result of direct marketing, and the Seller enters into this Agreement relying entirely upon such a warranty by the Purchaser**.",
            "initial": False,
        },
        {
            "type": 2,
            "section": "32.4",
            "text": "In the event of this Agreement being subject to the provisions of Section 16 read with Section 20 (2) (a) of the aforementioned Act, **it shall be deemed to be a resolutive condition to the effect that the Seller shall be entitled, in his sole discretion, to declare this Agreement null and void**, whereafter all amounts paid by the Purchaser, will be refunded and possession and occupation of the Property, will be returned to the Seller.",
            "initial": True,
        },
        {
            "type": 2,
            "section": "32.5",
            "text": "Kindly ensure that before signing this Agreement that you have had an adequate opportunity to understand these terms.  If you do not understand these terms or if you do not appreciate their effect, please ask for an explanation and do not sign the Agreement until the terms have been explained to your satisfaction.",
            "initial": False,
        },
        {
            "type": 2,
            "section": "32.6",
            "text": "The Sale of this property therefore constitutes a 'special-order' as may be contemplated by Section 17 of the Consumer Protection Act No. 68 of 2008",
            "initial": True,
        },
        {
            "type": 1,
            "section": "33.",
            "text": "**MORA INTEREST BREACH BY PURCHASER**",
            "initial": False,
        },
        {
            "type": 2,
            "section": "33.1",
            "text": "**Should the Purchaser commit any of the following mora interest breaches in terms of this Agreement**, and/or fails to comply with any of the provisions hereof, namely:",
            "initial": False,
        },
        {
            "type": 3,
            "section": "33.1.1",
            "text": "failure to pay the balance purchase price as set out in clause **E3** of the Information Schedule, in the stipulated timeframe set out in clause 4.5 of this Agreement;",
            "initial": False,
        },
        {
            "type": 3,
            "section": "33.1.2",
            "text": "failure to attend to the signing of the required transfer and mortgage loan documents (if any) in the timeframes stipulated in clause 9.7 of this Agreement;",
            "initial": False,
        },
        {
            "type": 2,
            "section": "",
            "text": "**then *ipso facto*, i.e. automatically from the 7th day until the date on which such breach is remedied, both days inclusive the Purchaser shall be in breach of the Agreement and will be liable to pay the Seller penalty interest on the full purchase price, in addition thereto,**",
            "initial": True,
        },
        {
            "type": 2,
            "section": "",
            "text": "the Seller shall be entitled to give the Purchaser 7 (seven) days' notice in writing to remedy such breach and/or failure, unless otherwise stated in this Agreement, and if the Purchaser fails to comply with such notice, then the Seller shall forthwith be entitled (but not obliged) without prejudice to any other rights or remedies which the Seller may have in law, including the right to claim damages:",
            "initial": False,
        },
        {
            "type": 3,
            "section": "(a)",
            "text": "to cancel this Agreement (in which event the Purchaser shall forfeit all monies paid to the Seller, its attorneys or its agent(s) in terms of this Agreement); or",
            "initial": False,
        },
        {
            "type": 3,
            "section": "(b)",
            "text": "to claim immediate performance and/or payment of all the obligations of the Purchaser in terms of this Agreement, including payment of unpaid balance of the purchase price; or",
            "initial": False,
        },
        {
            "type": 3,
            "section": "(c)",
            "text": "to claim mora interest as set out in the provisions of Clause 7.2.1 of this Agreement.",
            "initial": False,
        },
        {
            "type": 2,
            "section": "33.2",
            "text": "Should the Seller take steps against the Purchaser pursuant to a breach by the Purchaser of this Agreement, then without prejudice to any other rights which the Seller may have, the Seller shall be entitled to recover from the Purchaser all legal costs incurred by it including attorney/client charges, tracing fees and such collection commission as the Seller is obliged to pay to its attorneys.",
            "initial": False,
        },
        {
            "type": 2,
            "section": "33.3",
            "text": "Without prejudice to all or any of the rights of the Seller in terms of this Agreement, should the Purchaser fail to pay any amount due by the Purchaser in terms of this Agreement on due date, then the Purchaser shall pay the Seller interest thereon at the prime rate plus 2% calculated from the due date for payment until the actual date of payment, both dates inclusive.",
            "initial": False,
        },
        {
            "type": 1,
            "section": "34.",
            "text": "**MARKETING BY SELLER**",
            "initial": False,
        },
        {
            "type": 1,
            "section": "",
            "text": "The Purchaser acknowledges that the marketing material used by the Seller, to illustrate the interior and/or exterior of the Unit/Development, are for illustrative and representative purposes only, and does it not in any way form part of this Agreement, nor the stated specifications in Annexures **B & G**.",
            "initial": False,
        },
        {
            "type": 1,
            "section": "35.",
            "text": "**SIGNATURES**",
            "initial": False,
        },
        {
            "type": 1,
            "section": "",
            "text": "It is recorded that this document is intended to be signed firstly by the Purchaser and thereafter by the Seller. The Purchaser acknowledges that his signature hereto constitutes an irrevocable offer by him for the purchase of the Property on the terms and conditions set out herein, which offer shall remain irrevocable until 17:00 on the 8th day commencing on date signed by the Purchaser for acceptance by the Seller at any time prior hereto.",
            "initial": False,
        },
        {
            "type": 1,
            "section": "36.",
            "text": "**PROTECTION OF PERSONAL INFORMATION ACT**",
            "initial": False,
        },
        {
            "type": 1,
            "section": "",
            "text": "The Purchaser hereby consents to the Seller and/or the Agent furnishing the Mortgage Originator, the Financial Institution granting the loan, or the Transferring Attorney, with all such Personal Information of the Purchaser, as may be required to process any loan application by the Purchaser, or to attend to the transfer process.",
            "initial": False,
        },
    ]
    return standard_conditions


