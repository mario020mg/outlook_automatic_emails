import win32com.client
ol=win32com.client.Dispatch("outlook.application")
olmailitem=0x0 #size of the new email


emailsto = []

for emailto in emailsto:
    newmail=ol.CreateItem(olmailitem)
    newmail.Subject= '‼ ¿PROBLEMAS DE STOCK? AMPLÍA TU FLOTA CON RENTING, AHORA DESDE 12 MESES ‼'
    newmail.To=emailto
    newmail.BCC='25633542@bcc.eu1.hubspot.com'
    
 #  Imágenes para introducir dentro del texto
 #  Voydriving logo
    attachment = newmail.Attachments.Add(r'C:\Users\Usuario\Desktop\Programación Outlook\voydrivinglogo.png')
    attachment.PropertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001F", "MyId1")
 #  Facebook logo
    attachment = newmail.Attachments.Add(r'C:\Users\Usuario\Desktop\Programación Outlook\Facebook.png')
    attachment.PropertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001F", "facebook")
#   Instagram logo
    attachment = newmail.Attachments.Add(r'C:\Users\Usuario\Desktop\Programación Outlook\Instagram.png')
    attachment.PropertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001F", "instagram")
#   Linkedin Logo
    attachment = newmail.Attachments.Add(r'C:\Users\Usuario\Desktop\Programación Outlook\Linkedin.png')
    attachment.PropertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001F", "linkedin")
    
    
    newmail.HTMLBody= r'''<p style='margin:0cm;margin-bottom:.0001pt;font-size:15px;font-family:"Calibri","sans-serif";text-align:center;'><span style='font-family:"Arial","sans-serif";'><img src= cid:MyId1 width="144" height="56"></span></p>
<p style='margin:0cm;margin-bottom:.0001pt;font-size:15px;font-family:"Calibri","sans-serif";'><span style='font-family:"Arial","sans-serif";'>Hola<span style="color:black;">,&nbsp;</span></span></p>
<p style='margin:0cm;margin-bottom:.0001pt;font-size:15px;font-family:"Calibri","sans-serif";'><span style='font-family:"Arial","sans-serif";color:black;'>&nbsp;</span></p>
<p style='margin:0cm;margin-bottom:.0001pt;font-size:15px;font-family:"Calibri","sans-serif";'><span style='font-family:"Arial","sans-serif";color:black;'>Despu&eacute;s de la tormenta, siempre llega la calma&hellip; &iquest;o no? Se acerca el verano, y con ello m&uacute;ltiples <strong>viajes, excursiones, un aqu&iacute; para all&aacute;&nbsp;</strong>y &iquest;qu&eacute; mejor, que todas esas personas contraten el veh&iacute;culo para hacer todo lo comentado anteriormente, qu&eacute; contigo?&nbsp;</span></p>
<p style='margin:0cm;margin-bottom:.0001pt;font-size:15px;font-family:"Calibri","sans-serif";'><span style='font-family:"Arial","sans-serif";color:black;'>&nbsp;</span></p>
<p style='margin:0cm;margin-bottom:.0001pt;font-size:15px;font-family:"Calibri","sans-serif";'><span style='font-family:"Arial","sans-serif";color:black;'>&iexcl;<strong>NOVEDAD!</strong>, te ofrecemos <u>ofertas de Renting para 12 meses,&nbsp;</u>&iquest;vas a dejar pasar esta oportunidad?&nbsp;</span></p>
<p style='margin:0cm;margin-bottom:.0001pt;font-size:15px;font-family:"Calibri","sans-serif";'><span style='font-family:"Arial","sans-serif";'>&nbsp;</span></p>
<p style='margin:0cm;margin-bottom:.0001pt;font-size:15px;font-family:"Calibri","sans-serif";'><span style='font-family:"Arial","sans-serif";color:black;'>En&nbsp;</span><strong><span style='font-family:"Arial","sans-serif";color:#1F4E79;'>VOY DRIVING&nbsp;</span></strong><span style='font-family:"Arial","sans-serif";color:black;'>contamos con un amplio abanico de ofertas en veh&iacute;culos de renting destinados a tu colectivo.&nbsp;</span></p>
<p style='margin:0cm;margin-bottom:.0001pt;font-size:15px;font-family:"Calibri","sans-serif";'><span style='font-family:"Arial","sans-serif";color:black;'>&nbsp;</span></p>
<div style='margin:0cm;margin-bottom:.0001pt;font-size:15px;font-family:"Calibri","sans-serif";'>
    <ul style="list-style-type: square;">
        <li><span style='font-family:"Arial","sans-serif";color:black;'><span style='color: rgb(0, 0, 0); font-family: Arial, "sans-serif"; font-size: 15px; font-style: normal; font-variant-ligatures: normal; font-variant-caps: normal; font-weight: 400; letter-spacing: normal; orphans: 2; text-align: left; text-indent: 0px; text-transform: none; white-space: normal; widows: 2; word-spacing: 0px; -webkit-text-stroke-width: 0px; background-color: rgb(255, 255, 255); text-decoration-thickness: initial; text-decoration-style: initial; text-decoration-color: initial; display: inline !important; float: none;'>&iquest;</span>Modelos? <strong>Compactos, SUV e industriales</strong>.&nbsp;</span></li>
    </ul>
</div>
<p style='margin-top:0cm;margin-right:0cm;margin-bottom:.0001pt;margin-left:36.0pt;font-size:15px;font-family:"Calibri","sans-serif";'><span style='font-family:"Arial","sans-serif";'>&nbsp;</span></p>
<p style='margin-top:0cm;margin-right:0cm;margin-bottom:.0001pt;margin-left:36.0pt;font-size:15px;font-family:"Calibri","sans-serif";'><span style='font-family:"Arial","sans-serif";'>&nbsp;</span></p>
<div style='margin:0cm;margin-bottom:.0001pt;font-size:15px;font-family:"Calibri","sans-serif";'>
    <ul style="margin-bottom: 0cm; list-style-type: square;">
        <li style='margin:0cm;margin-bottom:.0001pt;font-size:15px;font-family:"Calibri","sans-serif";'><span style='font-family:"Arial","sans-serif";'>&iquest;Disponibilidad? Hasta fin de stock&nbsp;</span><strong><span style='font-family:"Segoe UI Symbol","sans-serif";color:#92D050;background:white;'>✔</span></strong><span style='font-family:"Arial","sans-serif";'>&nbsp;</span></li>
    </ul>
</div>
<p style='margin-top:0cm;margin-right:0cm;margin-bottom:.0001pt;margin-left:36.0pt;font-size:15px;font-family:"Calibri","sans-serif";'><span style='font-family:"Arial","sans-serif";'>&nbsp;</span></p>
<p style='margin-top:0cm;margin-right:0cm;margin-bottom:.0001pt;margin-left:36.0pt;font-size:15px;font-family:"Calibri","sans-serif";'><span style='font-family:"Arial","sans-serif";'>&nbsp;</span></p>
<p style='margin-top:0cm;margin-right:0cm;margin-bottom:.0001pt;margin-left:36.0pt;font-size:15px;font-family:"Calibri","sans-serif";'><span style='font-family:"Arial","sans-serif";color:black;'>&nbsp;</span></p>
<div style='margin:0cm;margin-bottom:.0001pt;font-size:15px;font-family:"Calibri","sans-serif";'>
    <ul style="margin-bottom: 0cm; list-style-type: square;">
        <li style='margin:0cm;margin-bottom:.0001pt;font-size:15px;font-family:"Calibri","sans-serif";'><strong><span style='font-family:"Arial","sans-serif";color:red;'>Si t&uacute; flota es grande</span></strong><span style='font-family:"Arial","sans-serif";color:black;'>, y est&aacute;s interesado en contratar 10 o m&aacute;s veh&iacute;culos, d&eacute;janos ayudarte con el seguro, 0 preocupaciones&nbsp;</span></li>
    </ul>
</div>
<p style='margin-top:0cm;margin-right:0cm;margin-bottom:.0001pt;margin-left:36.0pt;font-size:15px;font-family:"Calibri","sans-serif";'><span style='font-family:"Arial","sans-serif";color:black;'>&nbsp;</span></p>
<p style='margin-top:0cm;margin-right:0cm;margin-bottom:.0001pt;margin-left:36.0pt;font-size:15px;font-family:"Calibri","sans-serif";'><span style='font-family:"Arial","sans-serif";'>&nbsp;</span></p>
<p style='margin:0cm;margin-bottom:.0001pt;font-size:15px;font-family:"Calibri","sans-serif";'><span style='font-family:"Arial","sans-serif";color:black;'>Adem&aacute;s, contar con veh&iacute;culos&nbsp;</span><span style='font-family:"Arial","sans-serif";'>de<span style="color:black;">&nbsp;renting en su flota cuenta con muchas ventajas, destacando, sobre todo las&nbsp;</span></span><span style='font-family:"Segoe UI Symbol","sans-serif";'>✅</span><span style='font-family:"Arial","sans-serif";color:black;'>&nbsp;<strong><u>CONTABLES&nbsp;</u>&nbsp; &nbsp;</strong></span><span style='font-family:"Segoe UI Symbol","sans-serif";'>y &nbsp;✅</span><span style='font-family:"Arial","sans-serif";color:black;'>&nbsp;<strong><u>FINANCIERAS&nbsp;</u>&nbsp;</strong></span></p>
<p style='margin:0cm;margin-bottom:.0001pt;font-size:15px;font-family:"Calibri","sans-serif";'><strong><span style='font-family:"Arial","sans-serif";color:black;'>&nbsp;</span></strong></p>
<p style='margin:0cm;margin-bottom:.0001pt;font-size:15px;font-family:"Calibri","sans-serif";'><strong><span style='font-size:21px;font-family:"Arial","sans-serif";color:#1F4E79;'>VOY DRIVING</span></strong><strong><span style='font-family:"Arial","sans-serif";color:#1F4E79;'>&nbsp;</span></strong><span style='font-family:"Arial","sans-serif";color:black;'>siempre busca facilitar la vida a sus clientes, por ello, nos esforzamos en buscar, siempre, lo mejor para vosotros <strong>&iquest;hablamos?</strong></span></p>
<p style='margin:0cm;margin-bottom:.0001pt;font-size:15px;font-family:"Calibri","sans-serif";'><span style='font-family:"Arial","sans-serif";color:black;'>&nbsp;</span></p>
<p style='margin:0cm;margin-bottom:.0001pt;font-size:15px;font-family:"Calibri","sans-serif";'><span style='font-family:"Arial","sans-serif";color:black;'>Saludos.</span></p>

<p style='margin:0cm;margin-bottom:.0001pt;font-size:15px;font-family:"Calibri","sans-serif";'><strong><span style="color:#1F497D;">Mario Manzanares Galva&ntilde;</span></strong></p>
<p style='margin:0cm;margin-bottom:.0001pt;font-size:15px;font-family:"Calibri","sans-serif";'><strong><span style="color:#1F497D;">Consultor&iacute;a &amp; Asesor&iacute;a Comercial &ndash; VOY DRIVING.</span></strong></p>
<p style='margin:0cm;margin-bottom:.0001pt;font-size:15px;font-family:"Calibri","sans-serif";'><strong><span style="color:#1F497D;">&nbsp;</span></strong></p>
<p style='margin:0cm;margin-bottom:.0001pt;font-size:15px;font-family:"Calibri","sans-serif";'><strong><span style="color:#1F497D;">Calle del Agua, n&ordm; 6 &ndash; Oficina 21.</span></strong></p>
<p style='margin:0cm;margin-bottom:.0001pt;font-size:15px;font-family:"Calibri","sans-serif";'><strong><span style="color:#1F497D;">03400 Villena, Alicante, Espa&ntilde;a.</span></strong></p>
<p style='margin:0cm;margin-bottom:.0001pt;font-size:15px;font-family:"Calibri","sans-serif";'><strong><span style="color:#1F497D;">Telf: 681 954 975&ndash; 966 941 400.</span></strong></p>
<p style='margin:0cm;margin-bottom:.0001pt;font-size:15px;font-family:"Calibri","sans-serif";'><strong><span style="color:#1F497D;">Horario:&nbsp;</span></strong></p>
<p style='margin:0cm;margin-bottom:.0001pt;font-size:15px;font-family:"Calibri","sans-serif";'><strong><span style="color:#1F497D;">L a J: 9:00h - 14:00h // 15:00 &ndash; 17:00h</span></strong></p>
<p style='margin:0cm;margin-bottom:.0001pt;font-size:15px;font-family:"Calibri","sans-serif";'><strong><span style="color:#1F497D;">V: 8:00h &ndash; 15:00h</span></strong></p>
<p style='margin:0cm;margin-bottom:.0001pt;font-size:15px;font-family:"Calibri","sans-serif";'><strong><span style="color:#1F497D;">E-mail: asesor@voydriving.com&nbsp;</span></strong></p>
<p style='margin:0cm;margin-bottom:.0001pt;font-size:15px;font-family:"Calibri","sans-serif";'><strong><span style="color:#1F497D;">Web:&nbsp;</span></strong><a href="https://voydriving.com/"><strong><span style="color:blue;">https://voydriving.com/</span></strong></a><strong><span style="color:#1F497D;">&nbsp;</span></strong></p>
<p style='margin:0cm;margin-bottom:.0001pt;font-size:15px;font-family:"Calibri","sans-serif";'><span style="color:#1F497D;"><img src= cid:MyId1   alt="VoyDriving">&nbsp;</span></p>
<p style='margin:0cm;margin-bottom:.0001pt;font-size:15px;font-family:"Calibri","sans-serif";'><strong><span style="color:#1F497D;">&nbsp;</span></strong></p>
<p style='margin:0cm;margin-bottom:.0001pt;font-size:15px;font-family:"Calibri","sans-serif";'><a href="https://www.facebook.com/voyrenting/"><strong><span style="color:blue;text-decoration:none;"> <img src= cid:facebook ></span></strong></a>
<a href="https://www.instagram.com/voyrenting/?hl=es"><strong><span style="color:blue;text-decoration:none;"> <img src= cid:instagram ></span></strong></a>
<a href="https://www.linkedin.com/company/voy-renting/?viewAsMember=true"><strong><span style="color:blue;text-decoration:none;"> <img src= cid:linkedin ></span></strong></a> </p>
<p style='margin:0cm;margin-bottom:.0001pt;font-size:15px;font-family:"Calibri","sans-serif";text-align:justify;'><strong><span style="font-size:10px;color:#1F497D;">AVISO LEGAL:</span></strong><span style="font-size:10px;color:#1F497D;">&nbsp;La informaci&oacute;n contenida en este correo electr&oacute;nico, y en su caso en los documentos adjuntos, es informaci&oacute;n privilegiada para uso exclusivo de la persona y/o personas a las que va dirigido. No est&aacute; permitido el acceso a este mensaje a cualquier otra persona distinta a los indicados. Si Usted no es uno de los destinatarios, cualquier duplicaci&oacute;n, reproducci&oacute;n, distribuci&oacute;n, as&iacute; como cualquier uso de la informaci&oacute;n contenida en &eacute;l o cualquiera otra acci&oacute;n u omisi&oacute;n tomada en relaci&oacute;n con el mismo, est&aacute; prohibida y puede ser ilegal. En dicho caso, por favor notif&iacute;quelo al remitente y proceda a la eliminaci&oacute;n de este correo electr&oacute;nico as&iacute; como de sus adjuntos si los hubiere.&nbsp;</span></p>
<p style='margin:0cm;margin-bottom:.0001pt;font-size:15px;font-family:"Calibri","sans-serif";text-align:justify;'><span style="font-size:10px;color:#1F497D;">&nbsp;</span></p>
<p style='margin:0cm;margin-bottom:.0001pt;font-size:15px;font-family:"Calibri","sans-serif";text-align:justify;'><span style="font-size:10px;color:#1F497D;">Asimismo y en cumplimiento de Ley Org&aacute;nica 3/2018 de protecci&oacute;n de datos de car&aacute;cter personal y garant&iacute;a de los derechos digitales y del Reglamento Europeo RGPD 679/2016 le informamos que sus datos est&aacute;n siendo objeto de tratamiento por parte de VOY WORLD, S.L.U.con NIF B42706309, con la finalidad del mantenimiento y gesti&oacute;n de relaciones comerciales y administrativas. La base jur&iacute;dica del tratamiento es la aplicaci&oacute;n de medidas precontractuales o ejecuci&oacute;n de un contrato en el que usted es parte. En caso de ser el contacto de una persona jur&iacute;dica o empresario individual, se basar&aacute; en el inter&eacute;s leg&iacute;timo. No se prev&eacute;n cesiones y/o transferencias internacionales de datos. Para ejercitar sus derechos puede dirigirse a VOY WORLD, S.L.U., domiciliada en PARTIDA PE&Ntilde;ARRUBIA, 104, &nbsp;03400, VILLENA(ALICANTE), o bien por email a&nbsp;</span><a href="mailto:hola@voydriving.com"><span style="font-size:10px;color:blue;">hola@voydriving.com</span></a><span style="font-size:10px;color:#1F497D;">, &nbsp;con el fin de ejercer sus derechos de acceso, rectificaci&oacute;n, supresi&oacute;n (derecho al olvido), limitaci&oacute;n de tratamiento, portabilidad de los datos, oposici&oacute;n, y a no ser objeto de decisiones automatizadas, indicando como Asunto: &ldquo;Derechos Ley Protecci&oacute;n de Datos&rdquo;. Asimismo, tiene derecho a presentar una reclamaci&oacute;n ante la Agencia Espa&ntilde;ola de Protecci&oacute;n de Datos.</span></p>

'''
    attach= r'C:\Users\Usuario\Desktop\Programación Outlook\02. 2023 VOY DRIVING - Línea Arena (exclusiva colectivos).pdf'
    newmail.Attachments.Add(attach)
# To display the mail before sending it
#   newmail.Display() 
    newmail.Send()