import pdfkit
from lxml import etree
import os
import sys

def resource_path(relative_path):
    """
    Get the absolute path to a resource, whether running as a script or as a PyInstaller bundle.
    """
    base_path = getattr(sys, '_MEIPASS', os.path.abspath("."))  # _MEIPASS is defined by PyInstaller during runtime.
    return os.path.join(base_path, relative_path)


def xml_to_html(xml_path, xsl_path, output_html_path):
    """
    Transform the XML using XSLT to produce a fully rendered HTML file.
    """
    try:
        # Parse the XML and XSL files
        xml = etree.parse(xml_path)
        xsl = etree.parse(xsl_path)

        # Perform the transformation
        transform = etree.XSLT(xsl)
        result_tree = transform(xml)

        # Write the HTML output to a file
        with open(output_html_path, "wb") as html_file:
            html_file.write(etree.tostring(result_tree, pretty_print=True, encoding="UTF-8"))
        
        print(f"HTML successfully generated: {output_html_path}")
        return output_html_path
    except Exception as e:
        print(f"Error transforming XML to HTML: {e}")
        return None

def html_to_pdf(html_path, output_pdf_path):
    """
    Convert the rendered HTML file to a PDF using wkhtmltopdf.
    """
    try:

        options = {
            'orientation': 'Landscape',
            'page-size': 'Letter',
            'margin-top': '5mm',
            'margin-bottom': '5mm',
            'margin-left': '5mm',
            'margin-right': '5mm',
            'zoom': 1
            }
        wkhtmltopdf_path = resource_path("bin/wkhtmltopdf.exe")
        config = pdfkit.configuration(wkhtmltopdf=wkhtmltopdf_path)
        
        pdfkit.from_file(html_path, output_pdf_path, options=options, configuration=config)
        print(f"PDF successfully created: {output_pdf_path}")
    except Exception as e:
        print(f"Error converting HTML to PDF: {e}")

def xml_to_pdf(xml_path, xsl_path, pdf_path):
    """
    Full workflow: transform XML to HTML, then convert HTML to PDF.
    """
    # Intermediate HTML file path
    html_path = "C:/temp/FTSViewer/Sample_Config.html"
    
    # Step 1: Convert XML to HTML
    generated_html_path = xml_to_html(xml_path, xsl_path, html_path)
    if not generated_html_path:
        print("Failed to generate HTML from XML and XSL.")
        return

    # Step 2: Convert HTML to PDF
    html_to_pdf(generated_html_path, pdf_path)


"""
for testing purposes only

# Example Usage
if __name__ == "__main__":
    # Input XML and XSL files
    xml_file = "C:\\Users\\BedardO\\OneDrive - EC-EC\\Documents\\source\\repos\\FTSConfigViewer\\Sample_Config.xml"
    xsl_file = "C:\\Users\\BedardO\\OneDrive - EC-EC\\Documents\\source\\repos\\FTSConfigViewer\\FTSConfigViewer.xsl"

    # Output PDF file
    pdf_file = "output.pdf"

    # Create PDF from XML and XSL
    xml_to_pdf(xml_file, xsl_file, pdf_file)
"""