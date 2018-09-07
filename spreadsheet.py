import gspread
from oauth2client.service_account import ServiceAccountCredentials
import pprint

import smtplib
from email.MIMEMultipart import MIMEMultipart
from email.MIMEText import MIMEText

# use creds to create a client to interact with the Google Drive API
scope = ['https://spreadsheets.google.com/feeds',
         'https://www.googleapis.com/auth/drive']
creds = ServiceAccountCredentials.from_json_keyfile_name('client_secret.json', scope)
client = gspread.authorize(creds)
sheet = client.open("Fichier import Api").sheet1

list_cells = sheet.findall('yes')

for x in list_cells:
	if sheet.cell(x.row, 9).value == 'no':
		sheet.update_cell(x.row, 10, sheet.cell(x.row, 6).value)
		sheet.update_cell(x.row, 9, 'yes')
		name = sheet.cell(x.row, 2).value
		email = sheet.cell(x.row, 4).value
		date = sheet.cell(x.row, 5).value
		fromaddr = "maisonsmontrealaises@gmail.com"
		toaddr = "{0}".format(email)
		msg = MIMEMultipart()
		msg['From'] = fromaddr
		msg['To'] = toaddr
		msg['Subject'] = "Bonjour {0}, votre bail arrive a sa fin le {1} !".format(name, date)
		html = """\
		<DOCTYPE! html>
		<head>
		<style>
		html {
			font-family: Arial, sans-serif;
		}
		.background {
			padding: 50;
		}
		.jumbotron {
			background-image: url("http://www.printawallpaper.com/upload/designs/white_triangles_stacked_detail.jpg");
			padding: 20px;
			margin-bottom: 10px;
			border-radius: 5px;
		}
		.display-4 {
			margin-left: 10px;
		}
		.lead {
			margin-left: 10px;
		}
		.shadow-lg {
		    box-shadow: 0 4px 8px 0 
		    	rgba(0, 0, 0, 0.2), 
		    	0 6px 20px 0 rgba(0, 0, 0, 0.19);
		    padding: 5px;
		    border-radius: 5px;
		}
		.font-weight-light {
			color: #045D87;
			margin-left: 10px;
		}
		</style>
		<meta charset="UTF-8">
		</head>
		<body>
			<header class="background">
				<div class="jumbotron">
			  		<h1 class="display-4"><b>Bonjour,</b></h1>
			  		<p class="lead">Vous recevez ce mail car votre bail fini dans <b>environ 30 jours.</b></p>
			  		<hr class="my-4">
			  		<p class="lead">Merci de bien vouloir nous envoyer un mail a <u>maisonmontrealaise@gmail.com</u> pour comfirmer votre depart ou bien le renouvellement de votre bail.</p>
				</div>
				<div class="shadow-lg">
					<h4 class="display-4">Informations suplementaires</h4>
					<p class="font-weight-light">
					maisonmontrealaise@gmail.com<br>
					514-999-6577<br>
					1672-1678 Rue Saint Christophe<br>
					Montreal (QC) H3L 2W8
				</div>
			</header>
		</body>
		</html>
		"""
		msg.attach(MIMEText(html, 'html'))
		server = smtplib.SMTP('smtp.gmail.com', 587)
		server.starttls()
		server.login(fromaddr, "MaisonMontrealArnoul!")
		text = msg.as_string()
		server.sendmail(fromaddr, toaddr, text)
		server.quit()