# patch gooxml to allow to search mergefields in the tables
cd .\vendor\baliance.com\gooxml\document\
Get-Content ..\..\..\..\patch\08d318b31de33e9febd4b9f1763bb5212aeed35b.patch | patch -u -b
