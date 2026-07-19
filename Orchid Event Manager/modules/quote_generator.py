from __future__ import annotations

from datetime import date
from pathlib import Path
import re

from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_RIGHT
from reportlab.lib.pagesizes import letter
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import inch
from reportlab.platypus import Image, Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle

PURPLE = colors.HexColor('#6827BD')
PURPLE_DARK = colors.HexColor('#2E125F')
LIGHT_PURPLE = colors.HexColor('#F5F0FB')
BORDER = colors.HexColor('#DDD2EA')
MUTED = colors.HexColor('#6F6877')


def _safe_filename(value: str) -> str:
    cleaned = re.sub(r'[^A-Za-z0-9._-]+', '_', str(value or '').strip()).strip('_')
    return cleaned or 'Customer'


def _money(value: object) -> str:
    try:
        return f'${float(value):,.2f}'
    except (TypeError, ValueError):
        return '$0.00'


def generate_quote_pdf(
    employee_totals: list[dict],
    reports_root: Path,
    *,
    event_name: str,
    customer_name: str,
    quote_number: str,
    notes: str = '',
    logo_path: Path | None = None,
) -> Path:
    """Create a customer-facing quote from the imported employee totals."""
    output_dir = Path(reports_root) / 'Customer Quotes'
    output_dir.mkdir(parents=True, exist_ok=True)
    output_path = output_dir / f'{_safe_filename(quote_number)}_{_safe_filename(customer_name)}.pdf'

    doc = SimpleDocTemplate(
        str(output_path), pagesize=letter,
        leftMargin=0.55 * inch, rightMargin=0.55 * inch,
        topMargin=0.45 * inch, bottomMargin=0.5 * inch,
    )
    styles = getSampleStyleSheet()
    title_style = ParagraphStyle(
        'QuoteTitle', parent=styles['Title'], fontName='Helvetica-Bold',
        fontSize=23, leading=27, textColor=PURPLE_DARK, alignment=TA_CENTER,
    )
    subtitle_style = ParagraphStyle(
        'QuoteSubtitle', parent=styles['Normal'], fontName='Helvetica',
        fontSize=10.5, leading=14, textColor=MUTED, alignment=TA_CENTER,
    )
    cell_style = ParagraphStyle(
        'QuoteCell', parent=styles['Normal'], fontName='Helvetica',
        fontSize=9.5, leading=12, textColor=colors.HexColor('#26212C'), alignment=TA_LEFT,
    )
    right_style = ParagraphStyle('QuoteRight', parent=cell_style, alignment=TA_RIGHT)

    story = []
    if logo_path and Path(logo_path).exists():
        try:
            logo = Image(str(logo_path), width=2.65 * inch, height=0.78 * inch)
            logo.hAlign = 'CENTER'
            story.extend([logo, Spacer(1, 4)])
        except Exception:
            pass
    story.extend([
        Paragraph('CUSTOMER QUOTE', title_style),
        Paragraph('Orchid Uniforms &amp; Apparel', subtitle_style),
        Spacer(1, 12),
    ])

    info = [
        [Paragraph('<b>Customer</b>', cell_style), Paragraph(customer_name, cell_style),
         Paragraph('<b>Quote Number</b>', cell_style), Paragraph(quote_number, right_style)],
        [Paragraph('<b>Event</b>', cell_style), Paragraph(event_name, cell_style),
         Paragraph('<b>Quote Date</b>', cell_style), Paragraph(date.today().strftime('%B %d, %Y'), right_style)],
    ]
    info_table = Table(info, colWidths=[0.85*inch, 2.9*inch, 1.05*inch, 1.55*inch])
    info_table.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,-1), LIGHT_PURPLE),
        ('BOX', (0,0), (-1,-1), 0.7, BORDER),
        ('INNERGRID', (0,0), (-1,-1), 0.35, BORDER),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ('LEFTPADDING', (0,0), (-1,-1), 7), ('RIGHTPADDING', (0,0), (-1,-1), 7),
        ('TOPPADDING', (0,0), (-1,-1), 7), ('BOTTOMPADDING', (0,0), (-1,-1), 7),
    ]))
    story.extend([info_table, Spacer(1, 16)])

    rows = [[
        Paragraph('<b>Employee</b>', cell_style),
        Paragraph('<b>Order Total</b>', right_style),
        Paragraph('<b>Discount</b>', right_style),
        Paragraph('<b>Total</b>', right_style),
    ]]
    grand_total = 0.0
    for record in employee_totals:
        employee = str(record.get('employee') or record.get('company') or 'Employee')
        order_numbers = str(record.get('order_numbers') or '').strip()
        employee_text = employee
        if order_numbers:
            employee_text += f'<br/><font size="8" color="#6F6877">{order_numbers}</font>'
        order_total = float(record.get('order_total', 0) or 0)
        discount = float(record.get('discount', 0) or 0)
        total = float(record.get('total', order_total - discount) or 0)
        grand_total += total
        rows.append([
            Paragraph(employee_text, cell_style), Paragraph(_money(order_total), right_style),
            Paragraph(_money(discount), right_style), Paragraph(_money(total), right_style),
        ])
    rows.append([
        Paragraph('<b>FINAL TOTAL</b>', right_style), '', '', Paragraph(f'<b>{_money(grand_total)}</b>', right_style)
    ])

    totals_table = Table(rows, colWidths=[3.75*inch, 1.0*inch, 1.0*inch, 1.05*inch], repeatRows=1)
    totals_table.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,0), PURPLE),
        ('TEXTCOLOR', (0,0), (-1,0), colors.white),
        ('BACKGROUND', (0,-1), (-1,-1), LIGHT_PURPLE),
        ('BOX', (0,0), (-1,-1), 0.7, BORDER),
        ('INNERGRID', (0,0), (-1,-2), 0.3, colors.HexColor('#E9E3EE')),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ('LEFTPADDING', (0,0), (-1,-1), 7), ('RIGHTPADDING', (0,0), (-1,-1), 7),
        ('TOPPADDING', (0,0), (-1,-1), 6), ('BOTTOMPADDING', (0,0), (-1,-1), 6),
        ('SPAN', (0,-1), (2,-1)),
    ]))
    story.append(totals_table)

    if notes.strip():
        story.extend([
            Spacer(1, 16),
            Paragraph('<b>Notes</b>', cell_style),
            Spacer(1, 4),
            Paragraph(notes.replace('\n', '<br/>'), cell_style),
        ])
    story.extend([
        Spacer(1, 14),
        Paragraph(
            'This quote is based on the currently imported employee orders and saved discounts. '
            'Changes to employee orders or discounts may require a revised quote.',
            subtitle_style,
        ),
    ])
    doc.build(story)
    return output_path
