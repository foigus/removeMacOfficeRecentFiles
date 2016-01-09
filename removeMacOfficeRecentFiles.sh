#!/bin/bash

#Preference file location
preferenceFile=~/Library/Preferences/com.microsoft.office.plist
preferenceDomain=`basename -s .plist "${preferenceFile}"`

#Testing
#preferenceFile=~/com.microsoft.office.plist
#preferenceDomain=~/com.microsoft.office

#What number to report progress by when deleting individual Office recent items
countOfficeFilesBy=25

#Here's PlistBuddy!
plistBuddyLocation=/usr/libexec/PlistBuddy

#Check if file exists
if [ ! -f "${preferenceFile}" ]
then
	echo Preference file does not exist.  Exiting.
	exit 1
fi

#Report file size of preference
echo Beginning size of preference file: `du -h "${preferenceFile}" | awk '{ print $1 }'`

#For Office 2008 each recently used file is a separate preference key
#This leaves the first 20 (i.e. the most recent 20) files available for those who rely on recent files
for office2008Program in MSWD XCEL PPT
do
	#Create a sorted array of the the preference key numbers of the recent files
	office2008RecentFileKeyNumbers=( $( defaults read com.microsoft.office | grep 'File Aliases\\\\\\\\'"${office2008Program}"'' | awk -F= '{ print $1 }' | sed 's/\\\\\\//g' | awk -F\\ '{ print $3 }' | sed 's/\"//g' | sed 's/'"${office2008Program}"'//g' | sort -g ) )
	
	#Check if we have more than 20 recent files
	if [ "${#office2008RecentFileKeyNumbers[@]}" -gt 19 ] #Arrays are zero indexed
	then
		#We have at least 20 recent files--loop through deleting them
		echo -n "Deleting ${office2008Program} 2008 recent file keys: 1..."
		
		#Initialize a counter for counting deleted recent items
		let "c=0"

		for i in "${office2008RecentFileKeyNumbers[@]}"
		do
			if [ "${i}" -gt 19 ] #Arrays are zero indexed
			then
				#We are at an index greater than 19 (an excessive recent file), delete the File Alias and Access Date keys
				defaults delete "${preferenceDomain}" "2008\File Aliases\\${office2008Program}${i}" 
				defaults delete "${preferenceDomain}" "2008\MRU Access Date\\${office2008Program}${i}"
				
				#Add one to the counter
				let "c=c+1"
				
				#Determine if we should let the user know progress
				let "moduloCountBy=$c % $countOfficeFilesBy"
				if [ "${moduloCountBy}" -eq 0 ]
				then
					#Reached a point we should acknowledge progress
					echo -n "${c}..."
				fi
			fi
		done
		
		#Done deleting excess recent files for an application.  Echo out the final count and reset the counter
		echo "${c}"
		let "c=0"
	else
		echo 20 or fewer recent "${office2008Program}" 2008 recent file keys
	fi
done

#For Office 2011, all recent files are kept under a single per-program key in an array of dicts
#Thus there are two ways of handling Office 2011 recent files below
#The CFPreferences-safe method:
#  - Will delete all recent files by deleting the key itself
#  - Deletes all recent files, which some may view as a problem
#The CFPreferences-unsafe method:
#  - Uses PlistBuddy to delete array entries
#  - Leaves the 20 most recent files
#Uncomment the version that is best for your environment

#Office 2011 CFPreferences-safe
#Will lose all recent files!
#########
#for office2011Program in MSWD XCEL PPT3
#do
#	if defaults read "${preferenceDomain}" "14\File MRU\\${office2011Program}" >/dev/null 2>&1
#	then
#		defaults delete "${preferenceDomain}" "14\File MRU\\${office2011Program}"
#		echo Deleted "${office2011Program}" 2011 recent files key
#	else
#		echo "${office2011Program}" 2011 recent files key not found
#	fi
#done
#########

#Office 2011 CFPreferences-unsafe
#This leaves the first 20 (i.e. the most recent 20) array entries available for those who rely on recent files
#########
##Synchronize CFPreferences before running PlistBuddy, hopefully cutting down on the chance PlistBuddy and CFPreferences collide
#/usr/bin/python <<EOF
#import CoreFoundation 
#
#CoreFoundation.CFPreferencesAppSynchronize("${preferenceDomain}")
#EOF
#
#for office2011Program in MSWD XCEL PPT3
#do
#	#Check if the preference key exists
#	if ! defaults read "${preferenceDomain}" "14\File MRU\\${office2011Program}" >/dev/null 2>&1
#	then
#		echo 20 or fewer recent "${office2011Program}" 2011 recent file array dicts
#	else
#		#Count the number of array items by counting the instances of "File Alias"
#		arrayItems=`defaults read "${preferenceDomain}" "14\File MRU\\\\${office2011Program}" | grep -c "File Alias"`
#		
#		#If there are more than 20 array items, let's delete the excess recent file array dicts
#		if [ "${arrayItems}" -gt 19 ] #PlistBuddy is zero indexed
#		then
#			#We have at least 20 recent file array dicts
#			echo -n "Deleting ${office2011Program} 2011 recent file array dicts: 1..."
#
#			#Initialize a counter
#			let "c=0"
#
#			#Deduct one from arrayItems due to going from a one indexed tool to a zero indexed tool
#			let "arrayItemsZeroIndex=arrayItems-1"
#
#			#Delete the recent file array dicts
#			for i in `seq "${arrayItemsZeroIndex}" 20`
#			do
#				"${plistBuddyLocation}" -c 'Delete :14\\File\ MRU\\'"${office2011Program}"':'"${i}"'' "${preferenceFile}"
#				let "c=c+1"
#				
#				#Determine if we should let the user know progress
#				let "moduloCountBy=$c % $countOfficeFilesBy"
#				if [ "${moduloCountBy}" -eq 0 ]
#				then
#					#Reached a point we should acknowledge progress
#					let "actualDeletedFiles=$i-20"
#					echo -n "${c}..."
#				fi			
#			done
#			echo "${c}"
#			
#			#Reset the counter
#			let "c=0"
#		fi
#	fi
#done
#
##Since this isn't CFPreferences safe
#killall cfprefsd
#########

#Report file size of preference
echo Final size of preference file: `du -h "${preferenceFile}" | awk '{ print $1 }'`

exit 0
