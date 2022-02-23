import win32com.client
import datetime, locale

locale.setlocale(locale.LC_ALL, 'pl_PL')
data = datetime.datetime.now()
data_out = (data.strftime("%B"))
data_filename = (data.strftime("%d %B %Y"))

outlook = win32com.client.Dispatch('outlook.application')
mail = outlook.CreateItem(0)
mail.To = ''
mail.Subject = 'IMEI Sender'
mail.HTMLBody = '<h3>IMEI Sender</h3>'
mail.Body = 'This message was created automatically by mail delivery software, if there is something wrong with the attachment please inform me, otherwise do not reply ;)'
mail.Attachments.Add('C:/Users/filip.krokos/Scripts/Project1/'+str(data_filename)+".xlsx")
mail.Send()



