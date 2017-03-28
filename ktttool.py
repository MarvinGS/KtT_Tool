# -*- coding: iso-8859-1 -*-
#!/usr/bin/env python

import cgi
import cgitb; cgitb.enable()  # for troubleshooting

from pyPdf import PdfFileWriter, PdfFileReader
import StringIO
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import mm

from reportlab.graphics.shapes import Drawing 
from reportlab.graphics.barcode.qr import QrCodeWidget 
from reportlab.graphics import renderPDF

from reportlab.platypus import Paragraph
from reportlab.lib.styles import getSampleStyleSheet

styles = getSampleStyleSheet()
styleN = styles['Normal']
styleN.fontSize=12
styleN.leading=15
styleN.textColor="blue"
import uuid


def pmenu(form):
    print """
    <p>
            Dieses Tool f&uuml;llt die Beitrittsunterlagen f&uuml;r den KtT e.V.
            aus und bettet einen QR-Code mit diesen Daten in das PDF ein. Damit erspart es
            dem Team das Entschl&uuml;sseln von Handschriften und wom&ouml;glich beim Abtippen
            auftretende Fehler. Dies funktioniert nat&uuml;rlich nur, wenn nach dem Generieren keine
            Änderungen mehr an den vorausgef&uuml;llten Daten vorgenommen
            werden! Falls doch noch Korrekturen vorgenommen werden, bitte deutlich markieren!
            <br/><br/>
            Die Daten werden nicht f&uuml;r andere Zwecke als die PDF-Generierung f&uuml;r den KtT
            verwendet und nicht an Dritte
            weitergegeben.
            <br/><br/>
            Nach dem Ausf&uuml;llen der nicht vorausgef&uuml;llten Felder und dem Unterschreiben bitte
            an die o.g. Anschrift senden oder beim KtT abgeben.
            <br/><br/>
            Bei Problemen mit diesem Tool oder Fragen dazu bitte an Patrick G&uuml;nther &uuml;ber
            ktttool@pcgi.de wenden.</p>
    
    <p><form method="post" action="#" class="pcgiform">
    <fieldset>
        <ol>
            <input type="hidden" name="target" value="poly"/>
	    <li><label for="Vorname">Vorname</label> <input type=text id="Vorname" name="Vorname" /></li>
	    <li><label for="Nachname">Nachname</label> <input type=text id="Nachname" name="Nachname" /></li>
	    <li><label for="Strasse">Stra&szlig;e, Nr.</label> <input id="Strasse" name="Strasse" /></li>
	    <li><label for="PLZOrt">PLZ und Ort</label> <input id="PLZOrt" name="PLZOrt" /></li>
	    <li><label for="Tel">Telefon</label> <input id="Tel" name="Tel" /></li>
	    <li><label for="Mail">E-Mail-Adresse</label> <input id="Mail" name="Mail" /></li>
	    <li><label for="Geb">Geburtsdatum</label> <input id="Geb" name="Geb"/></li>
	    <li><label for="Mitgliedsnummer">Mitgliedsnummer</label> <input id="Mitgliedsnummer" name="Mitgliedsnummer"/></li>
	    <li><label for="Mitgliedschaft">Mitgliedschaft</label> <select id="Mitgliedschaft" name="Mitgliedschaft"/><option>Voll</option><option>Erm</option></select></li>
	    <li><label for="Zusatzbeitrag">Zusatzbeitrag</label> <input id="Zusatzbeitrag" name="Zusatzbeitrag"/></li>
	    <li><label for="Spendenbesch">Spendenbesch</label> <input type="checkbox" id="Spendenbesch" value="JA" name="Spendenbesch"/></li>
	    <li><label for="Udat">Datum, Ort (der Unterschrift)</label> <input id="Udat" name="Udat"/></li>
	    <li><label for="Kontoinhaber">Kontoinhaber</label> <input id="Kontoinhaber" name="Kontoinhaber"/></li>
	    <li><label for="Bank">Bank</label> <input id="Bank" name="Bank"/></li>
	    <li><label for="IBAN">IBAN</label> <input id="IBAN" name="IBAN"/></li>
	    <li><label for="BIC">BIC</label> <input id="BIC" name="BIC"/></li>
	</ol>
    </fieldset>
    <p><input type="submit" value="PDF generieren"></p>
    </form></p>

    """

def puttext(can, Absatz, Breite, Hoehe, x, y):
    can.saveState()
    P = Paragraph(Absatz,styleN)
    w, h = P.wrap(Breite,Hoehe)
    P.drawOn(can,x,y)
    can.restoreState()

def ppoly(form):
    print """
  
    
    <p>Bitte warten...</p>

    <p>

    """

    Vorname = unicode(form.getvalue("Vorname", " "),"iso-8859-1")
    Nachname = unicode(form.getvalue("Nachname", " "),"iso-8859-1")
    Strasse = unicode(form.getvalue("Strasse", " "),"iso-8859-1")
    PLZOrt = unicode(form.getvalue("PLZOrt", " "),"iso-8859-1")
    Tel = unicode(form.getvalue("Tel", " "),"iso-8859-1")
    Mail = unicode(form.getvalue("Mail", " "),"iso-8859-1")
    Geb = unicode(form.getvalue("Geb", " "),"iso-8859-1")
    Mitgliedsnummer = unicode(form.getvalue("Mitgliedsnummer", " "),"iso-8859-1")
    Mitgliedschaft = unicode(form.getvalue("Mitgliedschaft", " "),"iso-8859-1")
    Zusatzbeitrag = unicode(form.getvalue("Zusatzbeitrag", " "),"iso-8859-1")
    Spendenbesch = unicode(form.getvalue("Spendenbesch", " "),"iso-8859-1")
    Udat = unicode(form.getvalue("Udat", " "),"iso-8859-1")
    Kontoinhaber = unicode(form.getvalue("Kontoinhaber", " "),"iso-8859-1")
    Bank = unicode(form.getvalue("Bank", " "),"iso-8859-1")
    IBAN = unicode(form.getvalue("IBAN", " "),"iso-8859-1")
    BIC = unicode(form.getvalue("BIC", " "),"iso-8859-1")

    print "</p>"

    qrw = QrCodeWidget(Vorname+";"+Nachname+";"+Strasse+";"+PLZOrt+";"+Tel+";"+Mail+";"+Geb+";"+Mitgliedsnummer+";"+Mitgliedschaft+";"+Zusatzbeitrag+";"+Mitgliedsnummer+";"+Spendenbesch+";"+Mitgliedsnummer+";"+Udat+";"+Kontoinhaber+";"+Bank+";"+IBAN+";"+BIC) 
    b = qrw.getBounds()

    w=b[2]-b[0] 
    h=b[3]-b[1] 

    d = Drawing(45,45,transform=[120./w,0,0,120./h,0,0]) 
    d.add(qrw)

    packet0 = StringIO.StringIO()
    # create a new PDF with Reportlab
    can = canvas.Canvas(packet0, pagesize=A4)

    leftmargin=52
    start=237
    line=9
    pos=start

    can.drawString(leftmargin * mm, pos * mm, Vorname)
    pos-=line
    can.drawString(leftmargin * mm, pos * mm, Nachname)
    pos-=line
    can.drawString(leftmargin * mm, pos * mm, Strasse)
    pos-=line
    can.drawString(leftmargin * mm, pos * mm, PLZOrt)
    pos-=line
    can.drawString(leftmargin * mm, pos * mm, Mail)
    pos-=line
    can.drawString(leftmargin * mm, pos * mm, Tel)
    pos-=line
    can.drawString(leftmargin * mm, pos * mm, Geb)
    pos-=line
    can.drawString(leftmargin * mm, pos * mm, Mitgliedsnummer)


    if Mitgliedschaft == "Voll":
        can.drawString(23 * mm, 153.5 * mm, "x")

    if Mitgliedschaft == "Erm":
        can.drawString(23 * mm, 148.3 * mm, "x")

    if Spendenbesch == "JA":
        can.drawString(23.5 * mm, 117.5 * mm, "x")


    can.drawString(102 * mm, 127 * mm, Zusatzbeitrag)
    can.drawString(21 * mm, 105 * mm, Udat)
    can.drawString(21 * mm, 33 * mm, Udat)
    can.drawString(51 * mm, 55.5 * mm, Kontoinhaber)
    can.drawString(134 * mm, 55.5 * mm, Bank)
    can.drawString(51 * mm, 46 * mm, IBAN)
    can.drawString(134 * mm, 46 * mm, BIC)

    renderPDF.draw(d, can, 150*mm, 204*mm)
    can.save()
    packet0.seek(0)
    new_p0 = PdfFileReader(packet0)





    # read your existing PDF
    existing_pdf = PdfFileReader(file("/var/www/vhosts/pcgi.de/tools.pcgi.de/membership_form.pdf", "rb"))
    output = PdfFileWriter()
    # add the "watermark" (which is the new pdf) on the existing page
    page = existing_pdf.getPage(0)
    page.mergePage(new_p0.getPage(0))
    output.addPage(page)

    filename="KtT-"+str(uuid.uuid4())

    # finally, write "output" to a real file
    outputStream = file("/var/www/vhosts/pcgi.de/tools.pcgi.de/output/"+filename+".pdf", "wb")
    output.write(outputStream)
    outputStream.close()

    print "<p>Fertig! "
    print "<a href='/output/"+filename+".pdf'>PDF herunterladen</a> (ggf. Rechtsklick + 'Speichern unter...')</p>"

def pproblem(form):
    print """
  
    
    <p>Problem!</p>

    """


def main():
    print "Content-Type: text/html;charset=iso-8859-1"
    print

    print """
    <html>

    <head><title>PCGi.de PDFTool V0.7 (Ausf&uuml;llhilfe f&uuml;r KtT-Mitgliedschaft)</title></head>
    <meta http-equiv="Content-Type" content="text/html;charset=iso-8859-1" />
    <link rel="stylesheet" type="text/css" media="screen" href="/css/form.css" />    
    <body style="font-family: sans-serif;">

    <h3> PCGi.de PDFTool V0.7 (Ausf&uuml;llhilfe f&uuml;r KtT-Mitgliedschaft)</h3>
    """
  
    form = cgi.FieldStorage()
    target = form.getvalue("target", "menu")

    if target == "menu":
        pmenu(form)
    elif target == "poly":
        ppoly(form)
    else:
        pproblem(form)
              
    print """
              </body>
              
              </html>
    """

main()
