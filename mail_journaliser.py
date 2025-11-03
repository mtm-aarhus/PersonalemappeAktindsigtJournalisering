import os
from datetime import datetime
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import mm
from reportlab.pdfgen import canvas

def save_communication_pdf(subject: str, to_email: str, from_email: str, body: str) -> str:
    """
    Gemmer en kommunikation som PDF i mappen 'journal_pdfs'.
    Returnerer den fulde sti til PDF-filen.
    """
    # Sørg for at der findes en mappe
    output_dir = os.path.join(os.getcwd(), "journal_pdfs")
    os.makedirs(output_dir, exist_ok=True)

    # Gør filnavnet sikkert
    safe_subject = "".join(ch for ch in subject if ch.isalnum() or ch in (" ", "_", "-")).strip()
    timestamp = datetime.now().strftime("%Y%m%d-%H%M%S")
    filename = f"{timestamp}_{safe_subject[:50] or 'kommunikation'}.pdf"
    filepath = os.path.join(output_dir, filename)

    # Lav PDF
    c = canvas.Canvas(filepath, pagesize=A4)
    width, height = A4
    margin = 20 * mm
    textobject = c.beginText(margin, height - margin)

    # Header
    textobject.setFont("Helvetica-Bold", 14)
    textobject.textLine(subject)
    textobject.moveCursor(0, 10)

    # Metadata
    textobject.setFont("Helvetica", 10)
    textobject.textLine(f"Til: {to_email}")
    textobject.textLine(f"Fra: {from_email}")
    textobject.textLine(f"Dato: {datetime.now().strftime('%d.%m.%Y %H:%M')}")
    textobject.moveCursor(0, 15)

    # Indhold
    textobject.setFont("Helvetica", 11)
    if body:
        for line in body.splitlines():
            textobject.textLine(line)

    c.drawText(textobject)
    c.showPage()
    c.save()

    return filepath

def save_application_pdf(subject: str, from_email: str, body: str, modtagelsesdato) -> str:
    """
    Gemmer anmodning og tekst
    Returnerer den fulde sti til PDF-filen.
    """
    # Sørg for at der findes en mappe
    output_dir = os.path.join(os.getcwd(), "journal_pdfs")
    os.makedirs(output_dir, exist_ok=True)
    
    #Formater dato
    modtagelsesdato = datetime.fromisoformat(modtagelsesdato).strftime("%d-%m-%Y %H:%M")
    # Gør filnavnet sikkert
    safe_subject = "".join(ch for ch in subject if ch.isalnum() or ch in (" ", "_", "-")).strip()
    filename = f"{modtagelsesdato}_{safe_subject[:50] or 'kommunikation'}.pdf"
    filepath = os.path.join(output_dir, filename)

    # Lav PDF
    c = canvas.Canvas(filepath, pagesize=A4)
    width, height = A4
    margin = 20 * mm
    textobject = c.beginText(margin, height - margin)

    # Header
    textobject.setFont("Helvetica-Bold", 14)
    textobject.textLine(subject)
    textobject.moveCursor(0, 10)

    # Metadata
    textobject.setFont("Helvetica", 10)
    textobject.textLine(f"Fra: {from_email}")
    textobject.textLine(f"Dato: {modtagelsesdato}")
    textobject.moveCursor(0, 15)

    # Indhold
    textobject.setFont("Helvetica", 11)
    if body:
        for line in body.splitlines():
            textobject.textLine(line)

    c.drawText(textobject)
    c.showPage()
    c.save()

    return filepath