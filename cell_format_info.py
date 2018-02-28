from lxml import etree
import os
import re
from string import ascii_letters
import webbrowser
import zipfile

# Simple tool for inspecting XLSX output.
# Generates a four-cell HTML table representing the format
# of a selected cell in a selected XLSX worksheet,
# using an XSL transformation of the SpreadsheetML file.


def generate_internal_xlsx_path(sheet_number):
    internal_path = "xl/worksheets/sheet%d.xml" % sheet_number
    return internal_path


def get_column_number(column_name):
    if not column_name:
        raise ValueError
    column_number = 0
    for letter in column_name:
        if letter not in ascii_letters:
            raise ValueError
        else:
            letter = letter.upper()
            letter_number = (ord(letter) - ord("A")) + 1
            column_number = (column_number * 26) + letter_number
    return column_number


def process_cell_id(cell_id):
    matched = re.match("([A-Z]+)([0-9]+)", cell_id)
    column = matched.group(1)
    row = matched.group(2)
    column_number = get_column_number(column)
    row_number = int(row)
    return column_number, row_number


def get_sheet_xml_from_xlsx(xlsx_path, sheet_number=1):
    unzipper = zipfile.ZipFile(xlsx_path)
    internal_path = generate_internal_xlsx_path(sheet_number)
    xml = unzipper.read(internal_path)
    return xml


def render_xslt(cell_id):
    row_number, column_number = process_cell_id(cell_id)
    xslt = xslt_for_XLSX % (cell_id, row_number, column_number)
    print cell_id
    return xslt


def apply_xslt_to_xml(xml, xslt):
    raw_xslt = etree.XML(xslt)
    prepped_xslt = etree.XSLT(raw_xslt)
    prepped_xml = etree.XML(xml)
    result_tree = prepped_xslt(prepped_xml)
    return result_tree


def main(save_path="ximple.html"):
    default_path = "New.xlsx"
    filepath = raw_input("Enter XLSX name or path (default %s): " % default_path)
    filepath = filepath.strip()
    sheet_number = raw_input("Enter worksheet number (default 1): ").strip()
    cell_id = raw_input("Enter cell name (e.g. A1): ").strip()
    if not filepath or not filepath.endswith("xlsx"):
        filepath = default_path
    try:
        sheet_number = int(sheet_number.strip())
    except ValueError:
        sheet_number = 1
    sheet_xml = get_sheet_xml_from_xlsx(filepath, sheet_number)
    xslt = render_xslt(cell_id)
    output_tree = apply_xslt_to_xml(sheet_xml, xslt)
    output_html = etree.tostring(output_tree)
    open(save_path, "w").write(output_html)
    webbrowser.open(save_path)


xslt_for_XLSX = '''\
<?xml version="1.0" encoding="UTF-8"?>
<xsl:stylesheet version="2.0"
xmlns:xsl="http://www.w3.org/1999/XSL/Transform" xpath-default-namespace="http://example.com">
    <xsl:variable name="cellid" select="%s"/>        
    <xsl:variable name="colnum" select="%d"/>
    <xsl:variable name="rownum" select="%d"/>
    <xsl:template match="/">
        <html>
        <h1>Format information for cell <xsl:value-of select="$cellid"/></h1>
        <table>
        <xsl:for-each select="//*[local-name() = 'col']">
            <xsl:if test="@min &lt;= $colnum and $colnum &lt;= @max">
              <tr>
              <td width="50%%">
              </td>
              <td width="50%%">
              <h3>Column <xsl:value-of select="$colnum"/></h3>
              <xsl:for-each select="@*">
                  <xsl:value-of select="concat(name(), ': ', .)"/>
                  <br/>
               </xsl:for-each>
               </td></tr>
            </xsl:if>
        </xsl:for-each>
        <tr>
        <td>
            <h3>Row <xsl:value-of select="$rownum"/></h3>
            <xsl:for-each select="//*[local-name() = 'row']">
                <br/>
                <xsl:if test="@r = $rownum">
                    <xsl:for-each select="@*">
                        <xsl:sort/>
                        <xsl:value-of select="concat(name(), ': ', .)"/>
                        <br/>
                    </xsl:for-each>
                </xsl:if>
            </xsl:for-each> 
        </td>
        <td>
        <h3>Cell <xsl:value-of select="$cellid"/></h3>
        </td>
        </tr>
        </table>
        </html>
    </xsl:template>
</xsl:stylesheet>'''


if __name__ == "__main__":
    main()
