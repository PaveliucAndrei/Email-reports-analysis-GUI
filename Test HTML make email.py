from win32com.client import Dispatch

OUTLOOK = Dispatch('Outlook.Application')
OUTLOOK_NameSpace = OUTLOOK.GetNameSpace('MAPI')

email = OUTLOOK.CreateItem(0)



email.HTMLBody = r"""<html>
<body>

<style>
p {
  line-height: 1;
}
</style>

<p>Good day,</p>
<p>The next SWIFTs are processed with errors:</p>
<p>MT940_20240404_070309_44811.txt_2024-04-04-07:12:55</p>
<p>MT940_20240404_070310_44825.txt_2024-04-04-07:13:46</p>
<p>MT940_20240404_070311_44826.txt_2024-04-04-07:14:02</p>
<p>MT940_20240404_070311_44831.txt_2024-04-04-07:13:36</p>
<p>MT940_20240410_070140_45123.txt_2024-04-10-07:04:49</p>
<p>MT940_20240410_070140_45128.txt_2024-04-10-07:05:10</p>
<p>MT940_20240628_070149_51058.txt_2024-06-28-07:08:27</p>
<p>MT940_20240628_070151_51089.txt_2024-06-28-07:10:41</p>
<p>The following interface deliveries were processed with errors:</p>
<p>20240409212134_D_AID279_SCD_DEPBAMMEOD_SPARKLN_20240409202433237337_290.csv_2024-04-09-23:03:21</p>
<p>20240409_IDV-Bestand_.csv_2024-04-09-22:38:22</p>
<p>BWDATEI_F_Sparkasse KÃ¶lnBonn,_20240322202110</p>
<p>Datenlieferung_HANSAINVEST_2024-04-02-08-09-20-20240402-081353536.xlsx_2024-04-02-08:35:35</p>
<p>Please check.</p>
<p>Best regards.</p>

</body>
</html>"""

email.display()