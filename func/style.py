from docx.shared import Cm, Pt

def setStyle(document):
    # Set font
    style = document.styles['Normal']
    style.font.name = 'TH SarabunPSK'
    style.paragraph_format.line_spacing = 1