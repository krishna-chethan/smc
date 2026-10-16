import io
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import nsdecls, qn

def set_cell_border(cell, **kwargs):
    """
    Set cell borders
    kwargs: top, bottom, left, right
    values: dict(sz=12, val='single', color='000000')
    """
    tcPr = cell._tc.get_or_add_tcPr()
    tcBorders = tcPr.first_child_found_in("w:tcBorders")
    if tcBorders is None:
        tcBorders = OxmlElement('w:tcBorders')
        tcPr.append(tcBorders)
    for edge in ('top', 'left', 'bottom', 'right', 'insideH', 'insideV'):
        edge_data = kwargs.get(edge)
        if edge_data:
            tag = 'w:{}'.format(edge)
            element = tcBorders.find(qn(tag))
            if element is None:
                element = OxmlElement(tag)
                tcBorders.append(element)
            for key, attr in [("val", "w:val"), ("color", "w:color"), ("sz", "w:sz"), ("space", "w:space")]:
                if key in edge_data:
                    element.set(qn(attr), str(edge_data[key]))

def generate_sparsh_docx(data):
    doc = Document()
    
    # Page setup - Margins 0.6 inch
    for section in doc.sections:
        section.top_margin = Inches(0.6)
        section.bottom_margin = Inches(0.6)
        section.left_margin = Inches(0.7)
        section.right_margin = Inches(0.7)

    # Helper styling functions
    def add_heading_1(text):
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = p.add_run(text)
        run.bold = True
        run.font.size = Pt(13)
        run.font.name = "Times New Roman"
        p.paragraph_format.space_after = Pt(2)
        return p

    def add_subheading(text):
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = p.add_run(text)
        run.bold = True
        run.font.size = Pt(11)
        run.font.name = "Times New Roman"
        p.paragraph_format.space_after = Pt(6)
        return p

    def add_body(text, bold_prefix="", space_after=4):
        p = doc.add_paragraph()
        p.paragraph_format.space_after = Pt(space_after)
        p.paragraph_format.line_spacing = 1.15
        if bold_prefix:
            r1 = p.add_run(bold_prefix)
            r1.bold = True
            r1.font.size = Pt(10.5)
            r1.font.name = "Times New Roman"
        r2 = p.add_run(text)
        r2.font.size = Pt(10.5)
        r2.font.name = "Times New Roman"
        return p

    # --- 1. PROCEEDINGS ---
    authority = data.get("procIssuedBy", "MEO")
    receipt_sign_by = data.get("receiptSignBy", authority)
    institution_type = data.get("institutionType", "MRC")
    school_office = data.get("procIssuedSchoolName", "") or data.get("schoolName", "")
    mandal = data.get("mandalName", "")
    office_title = f"{authority}, {school_office}, {mandal}".strip(", ")
    
    add_heading_1(f"PROCEEDINGS OF THE {office_title}".upper())
    ddo_name = data.get("ddoName", "")
    ddo_qual = data.get("ddoQualification", "")
    ddo_str = f"Present : {ddo_name}" + (f", {ddo_qual}" if ddo_qual else "")
    add_subheading(ddo_str)

    p_meta = doc.add_paragraph()
    p_meta.paragraph_format.space_after = Pt(6)
    r_prc = p_meta.add_run(f"Proceedings No: {data.get('prcNo', '')}")
    r_prc.bold = True
    r_prc.font.size = Pt(10.5)
    r_prc.font.name = "Times New Roman"
    
    # tab to right for date
    p_meta.add_run(f"\t\t\t\t\t\tDate: {data.get('proDate', '')}").bold = True

    # Subject & Reference
    add_body(f" {data.get('procSubject', f'PMSHRI - SNA SPARSH - {school_office} Funds - Sanction of expenditure in Financial Year {data.get('financialYear')} towards {data.get('purpose')} under {data.get('budgetReleasedUnder')} - Orders Issued - Reg.')}", bold_prefix="Sub: ")
    
    # Dynamic references
    raw_refs = data.get("references", [])
    if not raw_refs and data.get("spdLetterNo"):
        raw_refs = [data.get("spdLetterNo")]
    if not raw_refs:
        ref_text = " 1. None"
    else:
        ref_lines = [f"{i+1}. {r}" for i, r in enumerate(raw_refs)]
        ref_text = " " + "\n".join(ref_lines)
    add_body(ref_text, bold_prefix="Ref: ")

    p_star = doc.add_paragraph("***")
    p_star.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_star.paragraph_format.space_after = Pt(4)

    # Order Text
    add_heading_1("ORDER")
    tot_amt = data.get("totalAmount", 0)
    amt_words = data.get("totalAmountWords", "")
    purpose = data.get("purpose", "")
    fy = data.get("financialYear", "2026-27")
    hoa = data.get("hoa", "")
    budget_cat = data.get("budgetReleasedUnder", "")

    order_text = (
        f"With reference to the subject and references cited above, sanction is hereby accorded for an amount of "
        f"Rs.{tot_amt:,.2f}/- ({amt_words}) towards the expenditure incurred for '{purpose}' "
        f"under {budget_cat} during the Financial Year {fy}."
    )
    add_body(order_text, space_after=8)

    # Component info
    p_comp = doc.add_paragraph()
    p_comp.paragraph_format.space_after = Pt(6)
    p_comp.add_run("Major Component: ").bold = True
    p_comp.add_run(f"{data.get('majorComponent', '')}   |   ")
    p_comp.add_run("Sub Component: ").bold = True
    p_comp.add_run(f"{data.get('subComponent', '')}")

    # Table of Vendors
    vendors = data.get("vendors", [])
    table = doc.add_table(rows=1, cols=5)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = False

    headers = ["S.No.", "Description / Purpose", "Amount (Rs.)", "Vendor / Beneficiary Name", "Bank, A/C & IFSC"]
    hdr_cells = table.rows[0].cells
    for i, title in enumerate(headers):
        hdr_cells[i].text = title
        p = hdr_cells[i].paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.runs[0].bold = True
        p.runs[0].font.size = Pt(9.5)
        set_cell_border(hdr_cells[i], top=dict(sz=4, val='single', color='000000'),
                                      bottom=dict(sz=4, val='single', color='000000'))

    for idx, v in enumerate(vendors, 1):
        row_cells = table.add_row().cells
        row_cells[0].text = str(idx)
        row_cells[1].text = v.get("description", purpose)
        row_cells[2].text = f"{float(v.get('amount', 0)):,.2f}"
        row_cells[3].text = v.get("vendorName", "")
        b_name = f"{v.get('bankName', '')}\n" if v.get('bankName') else ""
        row_cells[4].text = f"{b_name}{v.get('accountNumber', '')}\nIFSC: {v.get('ifscCode', '')}"
        for c in row_cells:
            for p in c.paragraphs:
                for r in p.runs:
                    r.font.size = Pt(9.5)
            set_cell_border(c, top=dict(sz=2, val='single', color='CCCCCC'),
                               bottom=dict(sz=2, val='single', color='CCCCCC'))

    # Total row
    tot_row = table.add_row().cells
    tot_row[0].text = ""
    tot_row[1].text = "TOTAL"
    tot_row[1].paragraphs[0].runs[0].bold = True
    tot_row[2].text = f"Rs.{tot_amt:,.2f}"
    tot_row[2].paragraphs[0].runs[0].bold = True
    tot_row[3].text = ""
    tot_row[4].text = ""

    # Treasury instruction
    add_body("The Treasury Officer / PAO is requested to honour the bill and credit the amount directly to the respective vendor bank accounts as per the above details.", space_after=12)
    add_body(f"Head of Account (HOA): {hoa}", bold_prefix="", space_after=16)

    # Signature block
    p_sig = doc.add_paragraph()
    p_sig.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    p_sig.paragraph_format.space_after = Pt(20)
    p_sig.add_run(f"\n\n\nSignature of the {authority}\n{school_office}, {mandal}").bold = True

    p_copy = doc.add_paragraph()
    p_copy.add_run("Copies to:\n1. The District Educational Officer & Accounts Officer.\n2. The Sub-Treasury Officer / PAO concerned.\n3. Office copy.").font.size = Pt(9.5)

    # --- 2. ADVANCE STAMPED RECEIPTS ---
    for idx, v in enumerate(vendors, 1):
        doc.add_page_break()
        add_heading_1("ADVANCE STAMPED RECEIPT")
        
        p_dt = doc.add_paragraph()
        p_dt.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        p_dt.add_run(f"Date: {data.get('proDate', '')}").bold = True

        v_amt = float(v.get("amount", 0))
        from num_to_words import num_to_words_inr
        v_amt_words = num_to_words_inr(v_amt)

        receipt_text = (
            f"Received with thanks from the Office of {school_office}, Mdl: {mandal}, the sum of "
            f"Rs.{v_amt:,.2f}/- ({v_amt_words}) towards the full and final settlement "
            f"against Credit Bill / Invoice for providing goods or services rendered towards '{v.get('description', purpose)}'."
        )
        add_body(receipt_text, space_after=10)
        add_body("The payment is acknowledged through online e-payment transfer/online treasury clearance system directly into our bank account as per the authorization details provided below:", space_after=10)

        add_heading_1("VENDOR BANK DETAILS FOR E-PAYMENT")
        add_body(f"1. Name of Beneficiary / Vendor: {v.get('vendorName', '')}")
        add_body(f"2. Bank Account Number: {v.get('accountNumber', '')}")
        add_body(f"3. Bank & Branch Name: {v.get('bankName', '') or 'As per IFSC'}")
        add_body(f"4. IFSC Code: {v.get('ifscCode', '')}")
        add_body(f"5. GSTIN (if applicable): {v.get('gstin', 'N/A')}", space_after=20)

        # Space directly above vendor signature line for stamp and sign
        p_v_sig = doc.add_paragraph()
        p_v_sig.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        p_v_sig.paragraph_format.space_before = Pt(36)
        p_v_sig.add_run(f"\n\nSignature of Vendor: ____________________\nName: {v.get('vendorName', '')}\nMobile / Phone: ____________________")
        
        # Border separator
        doc.add_paragraph("─" * 55).alignment = WD_ALIGN_PARAGRAPH.CENTER
        
        # Counter signature
        add_heading_1(f"COUNTER SIGNATURE OF {receipt_sign_by}")
        cs_text = (
            "Certified that the services/materials mentioned above have been received in good condition, "
            "entered in the Stock Register and found correct. The payment may be sanctioned and disbursed online "
            "directly to the vendor's bank account."
        )
        add_body(cs_text, space_after=25)
        
        p_cs_sig = doc.add_paragraph()
        p_cs_sig.add_run(f"Date: {data.get('proDate', '')}\nPlace: {mandal}\nUDISE Code: {data.get('udiseCode', '')}")
        p_cs_sig.add_run(f"\t\t\t\tSignature of {receipt_sign_by} with Office Seal").bold = True

    # --- 3. NON-PAYMENT CERTIFICATE ---
    doc.add_page_break()
    add_heading_1("NON-PAYMENT CERTIFICATE")
    p_dt = doc.add_paragraph()
    p_dt.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    p_dt.add_run(f"Date: {data.get('proDate', '')}").bold = True

    npc_text = (
        f"Certified that an amount of Rs.{tot_amt:,.2f}/- ({amt_words}) has not been drawn and paid previously. "
        f"The expenditure is claimed for the first time under the Head of Account: {hoa} ({budget_cat}) "
        f"towards '{purpose}' for the Financial Year {fy}."
    )
    add_body(npc_text, space_after=30)
    
    p_npc_sig = doc.add_paragraph()
    p_npc_sig.add_run(f"Date: {data.get('proDate', '')}\nPlace: {mandal}")
    p_npc_sig.add_run(f"\t\t\t\t\tSignature of {authority}\n\t\t\t\t\tOffice Seal").bold = True

    # --- 4. COMMITTEE RESOLUTION (School mode only) ---
    if institution_type == "School":
        doc.add_page_break()
        add_heading_1(f"EXECUTIVE RESOLUTION – {purpose.upper()}")
        add_subheading(f"School / Institution: {school_office}, Mandal: {mandal} | UDISE: {data.get('udiseCode', '')}")

        p_res_date = doc.add_paragraph()
        p_res_date.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        p_res_date.add_run(f"Meeting Date: {data.get('resolutionDate', '')}").bold = True

        res_p1 = (
            f"The School Management Committee / AAPC Chairperson, {data.get('schoolIncharge', authority)} and staff of "
            f"{school_office}, Mandal: {mandal} met on {data.get('resolutionDate', '')} and unanimously resolved to utilize the grant "
            f"of Rs.{tot_amt:,.2f}/- ({amt_words}) received under SNA SPARSH for FY {fy} strictly as per Government guidelines."
        )
        add_body(res_p1, space_after=10)

        res_p2 = (
            f"Out of the sanctioned amount, Rs.{tot_amt:,.2f}/- ({amt_words}) shall be paid directly to the concerned vendor(s) "
            f"through the IFMIS / SNA SPARSH portal. The Verification Committee has thoroughly verified the bills, invoices, "
            f"and supplied items/services, and recommends the final sanction and clearance."
        )
        add_body(res_p2, space_after=10)
        add_body("It is resolved that the grant has been utilized solely for approved institutional requirements and scheme objectives.", space_after=20)

        # Signatures
        add_heading_1("MEMBERS & STAFF PRESENT")
        st_tbl = doc.add_table(rows=5, cols=3)
        st_tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
        for r_i in range(5):
            st_tbl.rows[r_i].cells[0].text = f"{r_i+1}. ____________________"
            st_tbl.rows[r_i].cells[1].text = f"{r_i+6}. ____________________"
            st_tbl.rows[r_i].cells[2].text = f"Signature: ________________"

        p_comm_sig = doc.add_paragraph()
        p_comm_sig.paragraph_format.space_before = Pt(30)
        p_comm_sig.add_run(f"Signature of SMC / AAPC Chairperson\t\t\t\tSignature of {authority}").bold = True

    # --- 5. UTILIZATION CERTIFICATE ---
    doc.add_page_break()
    add_heading_1("UTILIZATION CERTIFICATE")
    add_subheading(f"FOR THE FINANCIAL YEAR {fy}")
    add_subheading(f"UNDER SNA SPARSH ACCOUNT – {purpose.upper()}")
    add_subheading(f"{school_office}, Mandal: {mandal} (UDISE: {data.get('udiseCode', '')})")

    uc_p1 = (
        f"Certified that a sum of Rs.{data.get('sanctionedAmount', tot_amt):,.2f}/- ({amt_words}) was sanctioned and received "
        f"under SNA SPARSH during the Financial Year {fy} towards the activity of '{purpose}'."
    )
    add_body(uc_p1, space_after=10)

    bal = float(data.get("sanctionedAmount', tot_amt", tot_amt)) - float(tot_amt)
    uc_p2 = (
        f"Out of this, an amount of Rs.{tot_amt:,.2f}/- ({amt_words}) has been utilized for the purpose for which it was "
        f"sanctioned as detailed category-wise. The unspent balance of Rs.{bal:,.2f}/- remains as on {data.get('proDate', '')}."
    )
    add_body(uc_p2, space_after=10)
    add_body("Certified that I have satisfied myself that the conditions on which the grant was sanctioned have been duly fulfilled and that I have exercised proper checks to ensure that the money was actually utilized for the purpose for which it was sanctioned.", space_after=25)

    p_uc_sig = doc.add_paragraph()
    p_uc_sig.add_run(f"Date: {data.get('proDate', '')}\nPlace: {mandal}")
    p_uc_sig.add_run(f"\t\t\t\t\tSignature of {authority}\n\t\t\t\t\tOffice Seal").bold = True

    target_stream = io.BytesIO()
    doc.save(target_stream)
    target_stream.seek(0)
    return target_stream
