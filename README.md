**A. With Skus Windows**
1. Copy skus to folder: C:\Windows\System32\spp\tokens\skus

2. Open CMD as Administrator. Run the following script:

set k1=VK7JG-NPHTM-C97JM-9MPGT-3V66T
cscript.exe %windir%\system32\slmgr.vbs /rilc
cscript.exe %windir%\system32\slmgr.vbs /upk >nul 2>&1
cscript.exe %windir%\system32\slmgr.vbs /ckms >nul 2>&1
cscript.exe %windir%\system32\slmgr.vbs /cpky >nul 2>&1
cscript.exe %windir%\system32\slmgr.vbs /ipk %k1%
sc config LicenseManager start= auto & net start LicenseManager
sc config wuauserv start= auto & net start wuauserv
clipup -v -o -altto c:\
start ms-settings:activation & exit

3. Replace key in the first line with the following suitable key 
Key Pro=VK7JG-NPHTM-C97JM-9MPGT-3V66T
Key Core=YTMG3-N6DKC-DKB77-7M9GH-8HVX7
Key CoreN=4CPRK-NM3K3-X6XXQ-RXX86-WXCHW	
Key ProN=2B87N-8KFHP-DKV6R-Y2C8J-PKCKT
Key CoreSL=BT79Q-G7N6G-PGBYW-4YWX6-6F4BT
Key Enterprise=XGVPP-NMH47-7TTHJ-W3FW7-8HV2C		
Key ProWorkstations=DXG7C-N36C4-C4HTG-X4T3X-2YV77
Key EnterpriseN=WGGHN-J84D6-QYCPR-T7PJ7-X766F
Key LTSB2015=CJ4JF-8KNY7-PCH2B-Y6B49-MG9R8
Key LTSB2016=NK96Y-D9CD8-W44CQ-R8YTK-DYJWX
Key LTSC2019=93MGM-NTFKD-6BK63-R6FYR-6Q9PB
Key Education=YNMGQ-8RYV3-4PGQ3-C8XTP-7CFBY
Key EducationN=84NGF-MHBT6-FXBX8-QWJK7-DRR8H
Key EnterpriseN=3V6Q6-NQXCX-V8YXR-9QCYV-QPFCT
Key ProfessionalEducation=8PTT6-RNW4C-6V7J2-C2D3X-MHBPB
Key ProfessionalEducationN=GJTYN-HDMQY-FRR76-HVGC7-QPF8P
Key ProfessionalSingleLanguage=G3KNM-CHG6T-R36X3-9QDG6-8M8K9
Key ProfessionalWorkstation=DXG7C-N36C4-C4HTG-X4T3X-2YV77
Key ProfessionalWorkstationN=WYPNQ-8C467-V2W6J-TX4WX-WT2RQ
Key IoTEnterprise=XQQYW-NFFMW-XJPBH-K8732-CKFFD

**B. With Skus Office**
== Office 2010
- Duong dan chua bo skus Office 2010
C:\Program Files\Microsoft Office\root\Licenses

Code for Retail
if exist "%ProgramFiles%\Microsoft Office\Office14\ospp.vbs" cd /d "%ProgramFiles%\Microsoft Office\Office14"
if exist "%ProgramFiles(x86)%\Microsoft Office\Office14\ospp.vbs" cd /d "%ProgramFiles(x86)%\Microsoft Office\Office14"

for /f %i in ('dir /b ..\root\Licenses\ProPlus_Retail*.xrm-ms') do cscript ospp.vbs /inslic:"..\root\Licenses\%i"
for /f %i in ('dir /b ..\root\Licenses\ProPlusMSDN_Retail*.xrm-ms') do cscript ospp.vbs /inslic:"..\root\Licenses\%i"

Code for Volume MAK
if exist "%ProgramFiles%\Microsoft Office\Office14\ospp.vbs" cd /d "%ProgramFiles%\Microsoft Office\Office14"
if exist "%ProgramFiles(x86)%\Microsoft Office\Office14\ospp.vbs" cd /d "%ProgramFiles(x86)%\Microsoft Office\Office14"

for /f %i in ('dir /b ..\root\Licenses\ProPlus_KMS_Client*.xrm-ms') do cscript ospp.vbs /inslic:"..\root\Licenses\%i"
for /f %i in ('dir /b ..\root\Licenses\ProPlus_MAK*.xrm-ms') do cscript ospp.vbs /inslic:"..\root\Licenses\%i"

-------------------------------------------------------------------------------------------------------------------------


== Office 2013

- Duong dan chua bo skus Office 2013
C:\Program Files\Microsoft Office\root\Licenses15

Code for Retail
if exist "%ProgramFiles%\Microsoft Office\Office15\ospp.vbs" cd /d "%ProgramFiles%\Microsoft Office\Office15"
if exist "%ProgramFiles(x86)%\Microsoft Office\Office15\ospp.vbs" cd /d "%ProgramFiles(x86)%\Microsoft Office\Office15"

for /f %i in ('dir /b ..\root\Licenses15\ProPlusr_Retail*.xrm-ms') do cscript ospp.vbs /inslic:"..\root\Licenses15\%i"
for /f %i in ('dir /b ..\root\Licenses15\ProPlusMSDNR_retail*.xrm-ms') do cscript ospp.vbs /inslic:"..\root\Licenses15\%i"

for /f %i in ('dir /b ..\root\Licenses15\VisioProR_Retail*.xrm-ms') do cscript ospp.vbs /inslic:"..\root\Licenses15\%i"
for /f %i in ('dir /b ..\root\Licenses15\VisioProMSDNR_Retail*.xrm-ms') do cscript ospp.vbs /inslic:"..\root\Licenses15\%i"

for /f %i in ('dir /b ..\root\Licenses15\ProjectProR_Retail*.xrm-ms') do cscript ospp.vbs /inslic:"..\root\Licenses15\%i"
for /f %i in ('dir /b ..\root\Licenses15\ProjectProMSDNR_Retail*.xrm-ms') do cscript ospp.vbs /inslic:"..\root\Licenses15\%i"

Code for Volume MAK
if exist "%ProgramFiles%\Microsoft Office\Office15\ospp.vbs" cd /d "%ProgramFiles%\Microsoft Office\Office15"
if exist "%ProgramFiles(x86)%\Microsoft Office\Office15\ospp.vbs" cd /d "%ProgramFiles(x86)%\Microsoft Office\Office15"

for /f %i in ('dir /b ..\root\Licenses15\ProPlusVL_MAK-*.xrm-ms') do cscript ospp.vbs /inslic:"..\root\Licenses15\%i"
for /f %i in ('dir /b ..\root\Licenses15\ProPlusVL_KMS_Client-*.xrm-ms') do cscript ospp.vbs /inslic:"..\root\Licenses15\%i"

for /f %i in ('dir /b ..\root\Licenses15\ProjectProVL_MAK-*.xrm-ms') do cscript ospp.vbs /inslic:"..\root\Licenses15\%i"

for /f %i in ('dir /b ..\root\Licenses15\VisioProVL_MAK-*.xrm-ms') do cscript ospp.vbs /inslic:"..\root\Licenses15\%i"

-----------------------------------------------------------------------------------------------------------------------------

== Office 2016

- Duong dan chua bo skus Office 2016
C:\Program Files\Microsoft Office\root\Licenses16

Code for Retail
if exist "%ProgramFiles%\Microsoft Office\Office16\ospp.vbs" cd /d "%ProgramFiles%\Microsoft Office\Office16"
if exist "%ProgramFiles(x86)%\Microsoft Office\Office16\ospp.vbs" cd /d "%ProgramFiles(x86)%\Microsoft Office\Office16"

for /f %i in ('dir /b ..\root\Licenses16\ProPlusR_Retail*.xrm-ms') do cscript ospp.vbs /inslic:"..\root\Licenses16\%i"
for /f %i in ('dir /b ..\root\Licenses16\ProPlusMSDNR_Retail*.xrm-ms') do cscript ospp.vbs /inslic:"..\root\Licenses16\%i"

for /f %i in ('dir /b ..\root\Licenses16\VisioProR_Retail*.xrm-ms') do cscript ospp.vbs /inslic:"..\root\Licenses16\%i"
for /f %i in ('dir /b ..\root\Licenses16\VisioProMSDNR_Retail*.xrm-ms') do cscript ospp.vbs /inslic:"..\root\Licenses16\%i"

for /f %i in ('dir /b ..\root\Licenses16\ProjectProR_Retail*.xrm-ms') do cscript ospp.vbs /inslic:"..\root\Licenses16\%i"
for /f %i in ('dir /b ..\root\Licenses16\ProjectProMSDNR_Retail*.xrm-ms') do cscript ospp.vbs /inslic:"..\root\Licenses16\%i"


Code for Volume MAK
if exist "%ProgramFiles%\Microsoft Office\Office16\ospp.vbs" cd /d "%ProgramFiles%\Microsoft Office\Office16"
if exist "%ProgramFiles(x86)%\Microsoft Office\Office16\ospp.vbs" cd /d "%ProgramFiles(x86)%\Microsoft Office\Office16"

for /f %i in ('dir /b ..\root\Licenses16\ProPlusVL_MAK-*.xrm-ms') do cscript ospp.vbs /inslic:"..\root\Licenses16\%i"
for /f %i in ('dir /b ..\root\Licenses16\ProPlusVL_KMS_Client-*.xrm-ms') do cscript ospp.vbs /inslic:"..\root\Licenses16\%i"

for /f %i in ('dir /b ..\root\Licenses16\VisioProVL_MAK-*.xrm-ms') do cscript ospp.vbs /inslic:"..\root\Licenses16\%i"
for /f %i in ('dir /b ..\root\Licenses16\VisioProVL_KMS_Client-*.xrm-ms') do cscript ospp.vbs /inslic:"..\root\Licenses16\%i"

for /f %i in ('dir /b ..\root\Licenses16\ProjectProVL_MAK-*.xrm-ms') do cscript ospp.vbs /inslic:"..\root\Licenses16\%i"
for /f %i in ('dir /b ..\root\Licenses16\ProjectProVL_KMS_Client-*.xrm-ms') do cscript ospp.vbs /inslic:"..\root\Licenses16\%i"

---------------------------------------------------------------------------------------------------------------------------------

== Office 2019

- Duong dan chua bo skus Office 2019
C:\Program Files\Microsoft Office\root\Licenses16

Code for Retail
if exist "%ProgramFiles%\Microsoft Office\Office16\ospp.vbs" cd /d "%ProgramFiles%\Microsoft Office\Office16"
if exist "%ProgramFiles(x86)%\Microsoft Office\Office16\ospp.vbs" cd /d "%ProgramFiles(x86)%\Microsoft Office\Office16"

for /f %i in ('dir /b ..\root\Licenses16\ProPlus2019R_Retail*.xrm-ms') do cscript ospp.vbs /inslic:"..\root\Licenses16\%i"
for /f %i in ('dir /b ..\root\Licenses16\ProPlus2019MSDNR_Retail*.xrm-ms') do cscript ospp.vbs /inslic:"..\root\Licenses16\%i"

for /f %i in ('dir /b ..\root\Licenses16\VisioPro2019R_Retail*.xrm-ms') do cscript ospp.vbs /inslic:"..\root\Licenses16\%i"
for /f %i in ('dir /b ..\root\Licenses16\VisioPro2019MSDNR_Retail*.xrm-ms') do cscript ospp.vbs /inslic:"..\root\Licenses16\%i"

for /f %i in ('dir /b ..\root\Licenses16\ProjectPro2019R_Retail*.xrm-ms') do cscript ospp.vbs /inslic:"..\root\Licenses16\%i"
for /f %i in ('dir /b ..\root\Licenses16\ProjectPro2019MSDNR_Retail*.xrm-ms') do cscript ospp.vbs /inslic:"..\root\Licenses16\%i"

Code for Volume MAK
if exist "%ProgramFiles%\Microsoft Office\Office16\ospp.vbs" cd /d "%ProgramFiles%\Microsoft Office\Office16"
if exist "%ProgramFiles(x86)%\Microsoft Office\Office16\ospp.vbs" cd /d "%ProgramFiles(x86)%\Microsoft Office\Office16"

for /f %i in ('dir /b ..\root\Licenses16\ProPlus2019VL_MAK_AE*.xrm-ms') do cscript ospp.vbs /inslic:"..\root\Licenses16\%i"
for /f %i in ('dir /b ..\root\Licenses16\ProPlus2019VL_KMS_Client_AE*.xrm-ms') do cscript ospp.vbs /inslic:"..\root\Licenses16\%i"

for /f %i in ('dir /b ..\root\Licenses16\VisioPro2019VL_MAK-*.xrm-ms') do cscript ospp.vbs /inslic:"..\root\Licenses16\%i"
for /f %i in ('dir /b ..\root\Licenses16\VisioPro2019VL_KMS_Client-*.xrm-ms') do cscript ospp.vbs /inslic:"..\root\Licenses16\%i"

for /f %i in ('dir /b ..\root\Licenses16\ProjectPro2019VL_MAK-*.xrm-ms') do cscript ospp.vbs /inslic:"..\root\Licenses16\%i"
for /f %i in ('dir /b ..\root\Licenses16\ProjectPro2019VL_KMS_Client-*.xrm-ms') do cscript ospp.vbs /inslic:"..\root\Licenses16\%i"

-------------------------------------------------------------------------------------------------------------------------------------

== Office 2021

- Duong dan chua bo skus Office 2021
C:\Program Files\Microsoft Office\root\Licenses16

Code for Retail
if exist "%ProgramFiles%\Microsoft Office\Office16\ospp.vbs" cd /d "%ProgramFiles%\Microsoft Office\Office16"
if exist "%ProgramFiles(x86)%\Microsoft Office\Office16\ospp.vbs" cd /d "%ProgramFiles(x86)%\Microsoft Office\Office16"

for /f %i in ('dir /b ..\root\Licenses16\ProPlus2021R_Retail*.xrm-ms') do cscript ospp.vbs /inslic:"..\root\Licenses16\%i"
for /f %i in ('dir /b ..\root\Licenses16\ProPlus2021MSDNR_Retail*.xrm-ms') do cscript ospp.vbs /inslic:"..\root\Licenses16\%i"

for /f %i in ('dir /b ..\root\Licenses16\VisioPro2021R_Retail*.xrm-ms') do cscript ospp.vbs /inslic:"..\root\Licenses16\%i"
for /f %i in ('dir /b ..\root\Licenses16\VisioPro2021MSDNR_Retail*.xrm-ms') do cscript ospp.vbs /inslic:"..\root\Licenses16\%i"

for /f %i in ('dir /b ..\root\Licenses16\ProjectPro2021R_Retail*.xrm-ms') do cscript ospp.vbs /inslic:"..\root\Licenses16\%i"
for /f %i in ('dir /b ..\root\Licenses16\ProjectPro2021MSDNR_Retail*.xrm-ms') do cscript ospp.vbs /inslic:"..\root\Licenses16\%i"

Code for Volume MAK
if exist "%ProgramFiles%\Microsoft Office\Office16\ospp.vbs" cd /d "%ProgramFiles%\Microsoft Office\Office16"
if exist "%ProgramFiles(x86)%\Microsoft Office\Office16\ospp.vbs" cd /d "%ProgramFiles(x86)%\Microsoft Office\Office16"

for /f %i in ('dir /b ..\root\Licenses16\ProPlus2021VL_MAK_AE*.xrm-ms') do cscript ospp.vbs /inslic:"..\root\Licenses16\%i"
for /f %i in ('dir /b ..\root\Licenses16\ProPlus2021VL_KMS_Client_AE*.xrm-ms') do cscript ospp.vbs /inslic:"..\root\Licenses16\%i"

for /f %i in ('dir /b ..\root\Licenses16\VisioPro2021VL_MAK*.xrm-ms') do cscript ospp.vbs /inslic:"..\root\Licenses16\%i"
for /f %i in ('dir /b ..\root\Licenses16\VisioPro2021VL_KMS_Client*.xrm-ms') do cscript ospp.vbs /inslic:"..\root\Licenses16\%i"

for /f %i in ('dir /b ..\root\Licenses16\ProjectPro2021VL_MAK*.xrm-ms') do cscript ospp.vbs /inslic:"..\root\Licenses16\%i"
for /f %i in ('dir /b ..\root\Licenses16\ProjectPro2021VL_KMS_Client*.xrm-ms') do cscript ospp.vbs /inslic:"..\root\Licenses16\%i"

-------------------------------------------------------------------------------------------------------------------------------------

== Office 2024

- Duong dan chua bo skus Office 2024
C:\Program Files\Microsoft Office\root\Licenses16

Code for Retail
if exist "%ProgramFiles%\Microsoft Office\Office16\ospp.vbs" cd /d "%ProgramFiles%\Microsoft Office\Office16"
if exist "%ProgramFiles(x86)%\Microsoft Office\Office16\ospp.vbs" cd /d "%ProgramFiles(x86)%\Microsoft Office\Office16"

for /f %i in ('dir /b ..\root\Licenses16\ProPlus2024R_Retail*.xrm-ms') do cscript ospp.vbs /inslic:"..\root\Licenses16\%i"

for /f %i in ('dir /b ..\root\Licenses16\VisioPro2024R_Retail*.xrm-ms') do cscript ospp.vbs /inslic:"..\root\Licenses16\%i"

for /f %i in ('dir /b ..\root\Licenses16\ProjectPro2024R_Retail*.xrm-ms') do cscript ospp.vbs /inslic:"..\root\Licenses16\%i"

Code for Volume MAK
if exist "%ProgramFiles%\Microsoft Office\Office16\ospp.vbs" cd /d "%ProgramFiles%\Microsoft Office\Office16"
if exist "%ProgramFiles(x86)%\Microsoft Office\Office16\ospp.vbs" cd /d "%ProgramFiles(x86)%\Microsoft Office\Office16"

for /f %i in ('dir /b ..\root\Licenses16\ProPlus2024VL_MAK_AE*.xrm-ms') do cscript ospp.vbs /inslic:"..\root\Licenses16\%i"
for /f %i in ('dir /b ..\root\Licenses16\ProPlus2024VL_KMS_Client_AE*.xrm-ms') do cscript ospp.vbs /inslic:"..\root\Licenses16\%i"

for /f %i in ('dir /b ..\root\Licenses16\VisioPro2024VL_MAK*.xrm-ms') do cscript ospp.vbs /inslic:"..\root\Licenses16\%i"
for /f %i in ('dir /b ..\root\Licenses16\VisioPro2024VL_KMS_Client*.xrm-ms') do cscript ospp.vbs /inslic:"..\root\Licenses16\%i"

for /f %i in ('dir /b ..\root\Licenses16\ProjectPro2024VL_MAK*.xrm-ms') do cscript ospp.vbs /inslic:"..\root\Licenses16\%i"
for /f %i in ('dir /b ..\root\Licenses16\ProjectPro2024VL_KMS_Client*.xrm-ms') do cscript ospp.vbs /inslic:"..\root\Licenses16\%i"

-------------------------------------------------------------------------------------------------------------------------------------

Ung dung them cac skus Office khac bang cach thay doi dia chi duong dan chua skus va ten skus muon nap them.

Office 2010 xoa (XX)
Office 2013 thay (XX) = 15
Office 2016, 2019, 2021, 2024 thay (XX) = 16

TEN_SKUS = Rename ten file SKUS, copy ten SKUS dan vao

for /f %i in ('dir /b ..\root\Licenses(xx)\TEN_SKUS*.xrm-ms') do cscript ospp.vbs /inslic:"..\root\Licenses(xx)\%i"
