# EWSAsposePDFDemo

This is a demo project to show the EWS nuget package(https://www.nuget.org/packages/Microsoft.Exchange.WebServices/) can be used to fetch the emails from the exchange box and to convert the mails fetched to PDF with attachments apended to it using Aspose PDF package (Free version - watermarked with evaluation copy - https://www.nuget.org/packages/Aspose.Pdf/) and Aspose Word package (Free version - watermarked with evaluation copy - https://www.nuget.org/packages/Aspose.Words/)

I have used a forked version of MSGReader project (https://github.com/seppu/MSGReader) froeked from https://github.com/Sicos1977/MSGReader to convert & format the EML data to memory stream object and to fetch the attachments. 

Please download this forked version of MSGReader (https://github.com/seppu/MSGReader) and place it in the same folder as ExchangePOC.sln as this project is also included in it. 
