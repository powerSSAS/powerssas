
set-alias installutil $env:windir\Microsoft.NET\Framework\v2.0.50727\installutil

cd \data\projects\powerSSAS\powerSSAS
copy bin\debug\powerSSAS.dll ..\powerSSAS.dll

installutil ..\powerSSAS.dll

add-pssnapin powerSSAS
# create a new drive
new-psdrive ssas powerSSAS localhost

cd ssas:
cd databases
cd "adventure works dw"

#remove-PSSnapin powerSSAS