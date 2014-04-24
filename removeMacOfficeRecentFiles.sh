#2011
#All recent files are kept under a single per-program key
defaults delete ~/Library/Preferences/com.microsoft.office "14\File MRU\XCEL"
defaults delete ~/Library/Preferences/com.microsoft.office "14\File MRU\PPT3"
defaults delete ~/Library/Preferences/com.microsoft.office "14\File MRU\MSWD"

#2008
#Each recent file in Office 2008 is a separate preference key
#This is good for 1500 recent files--the most I have seen is around 1000
for i in {1..1500}
do
	defaults delete ~/Library/Preferences/com.microsoft.office "2008\File Aliases\XCEL$i"
	defaults delete ~/Library/Preferences/com.microsoft.office "2008\MRU Access Date\XCEL$i"

	defaults delete ~/Library/Preferences/com.microsoft.office "2008\File Aliases\MSWD$i"
	defaults delete ~/Library/Preferences/com.microsoft.office "2008\MRU Access Date\MSWD$i"

	defaults delete ~/Library/Preferences/com.microsoft.office "2008\File Aliases\PPT$i"
	defaults delete ~/Library/Preferences/com.microsoft.office "2008\MRU Access Date\PPT$i"
done
