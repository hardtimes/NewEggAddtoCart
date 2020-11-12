#
#  This script will parse newegg for $searchCriteria, and if stock is found, open Chrome web broswer and attempt to add it to your cart.
#
#  Requirements:
#    - You have run the command "set-executionpolicy unrestricted" in powershell.  Run "Set-executionpolicy restricted" to reverse this when done with the script.
#    - Windows 10 - Tested on Windows 10 build 2004(spring 2020 release)
#    -Chrome 64 bit be installed(you can try with firefox, IE, edge, 32bit browsers, etc if you know what you are doing.  Testing only done with Chrome)  
#    ---You should have Chrome already open at a newegg.com at the cart when running this, to delay load times when an item is found in-stock.
#    -Open Internet Explorer and finsh the auto configuration prompt if you have not.  If you do not, the native powershell cmdlet invoke-webrequest will not work.
#    
#
#    **This WILL NOT log you into newegg or check you out.  You have to do those before and after. 
#
#  How to use:
#  -Open this file in "Windows Powershell ISE", which is in your start menu by default.  Search for it.
#  -Review and modify ##OPTIONS## below before running. 
#
# Written by: Paul Dickson

####### OPTIONS ########


# Change this as desired.
$searchCriteria="rtx 3080"
#$searchCriteria="rtx 3090"
#$searchCriteria="Anker Soundcore Liberty"  # for testing
#$searchCriteria="synology tryhi" # for testing

# Items that should be ignored if found in the search results.  Uses wildcard(*) matching
$searchIgnores=@(
    "ABS Gladiato*", # prebuilt pcs
    "CLX*", # prebuilt pcs
    #"Intel Core*" # intel core combos
    "Montech*" # just cases that matches the rtx 3080 search
)

# show ignored search items in search results.  They will list as ignored, and add to cart will not happen.
$showIgnores=$true
#$showIgnores=$false

# If set to true, the script will only show front page stock status, but will not confirm or try to add to cart.    
#  Good for testing new search criteria.  Be careful about re-testing every few seconds.  Even that will get you blocked on newegg.
#$testing=$true
$testing=$false

# Set this to $true to immediately stop the script after the first item found in stock and attempted to add to cart.
# Set to $false to continue checking and adding to cart when found.  Not recommended as it will likely get in the way of you checking out if it finds new items in stock by opening new Chrome tabs when it tries to add then to cart. Are you feeling lucky punk?
#$stopInStockOne=$true
$stopInStockOne=$false

# $stopInstockOne above must not be set to $true for this to have any effect
# If set to $true, this will stop the script after all currently instock items on the current check are attempted to be added to cart.
## ! Please consider canceling any items you purchase that are not the one you intend to keep for personal use.  Fuck scalpers.
$stopInStockAll=$true
#$stopInStockAll=$false

# delay in seconds between checks.  Without a delay you will be blocked from querying the site.  
# KEEP IN MIND people have been baned for refreshing newegg too frequently, causing all lookups to fail from the script.
$delay=15

# This is the default installation location for Chrome x64. Update as needed.
$chromeLocation="c:\Program Files\Google\Chrome\Application\chrome.exe"
#$chromeLocation="c:\SomeOtherFolder\chrome.exe"

# Set to $true to open the newegg shopping cart in chrome when the script first starts, for caching purposes.  Otherwise set to $false.
#$preloadShoppingCart=$true
$preloadShoppingCart=$false

# Set to newegg.ca, newegg.whatever as needed.  Only tested on newegg.com. Should work on others except:Does not work on neweggbusiness.com
$neweggDomain="newegg.com"

# Play a sound if item is found in stock
$playAlertOnStock=$true

# Number of times to play alert sound. 0 keeps playing until you stop it.
##  Not advisable to set to 0 if $stopInStockOne and $stopInStockAll are both $false, because then it will never stop playing.
$playAlertTimes=5

# Set to true if you want to do telegram alerts.  You MUST setup your own telegram bot. See https://www.itdroplets.com/automating-telegram-messages-with-powershell/ to setup a telegram bot.
#$doTelegram=$true
$doTelegram=$false

# Token and chat ID from your bot.  See https://www.itdroplets.com/automating-telegram-messages-with-powershell/ to setup a telegram bot.
#$TG_BOT_TOKEN = "YOURTOKENHERE"
#$TG_BOT_CHAT_ID = 1234567890

####### END OPTIONS ########


# open shopping cart in web browser to get it cached.  
if ($preloadShoppingCart){
    start-process $chromeLocation "https://secure.$neweggDomain/Shopping/ShoppingCart.aspx?Submit=view"
}

# Audio alert
function invoke-alert {
    param(
    $times=1,
    $wav="C:\windows\media\Windows Exclamation.wav"
    )

    $PlayWav=New-Object System.Media.SoundPlayer
    $PlayWav.SoundLocation=$wav
    
    $block={
        $times=$args[0]
        $PlayWav=New-Object System.Media.SoundPlayer
        $PlayWav.SoundLocation=$args[1]
        if ($times -ne 0){
            for ($i=1; $i -le $times; $i++){
                $PlayWav.playsync()
            }
        } else {
            while ($true){$PlayWav.playsync(); sleep 1}
        }

    }

    $job=start-job -ScriptBlock $block -ArgumentList $times,$wav
    $job
}


function stop-alert {
    if ([bool]($soundJob.PSobject.Properties.name -contains "id")){
        get-job $soundJob.id | Stop-Job 
    }
}


$itemsFound=$null
$doStop=$false
while (-not($doStop)){
    try{
        # search for all matching products stocked by newegg, regardless of "in stock" filter result on website
        #$search=Invoke-WebRequest "https://www.$neweggDomain/p/pl?d=$searchCriteria&N=4841" -UseBasicParsing -ea stop

        # search only for products the new egg website 'in stock' filter returns, and which also stocked by newegg. 
        $search=Invoke-WebRequest "https://www.$neweggDomain/p/pl?d=$searchCriteria&N=4131&N=4841" -UseBasicParsing -ea stop
        
        [psobject]$searchResults=$search
        $urls=$searchResults.links | ? href -match .*$searchCriteria.* | ? href -match .*\-Product | ? href -notmatch .*scrollFullinfo.* | select  -ExpandProperty href | sort -Unique
        $rawcontent=$searchResults.RawContent
    }catch{
        $_
        write-host "Could not parse search results. Retrying after the configured delay: $delay secs"
        $search
        sleep $delay
        continue
    }
    write-host "`n`nProducts found matching `"$($searchCriteria)`": $(($urls| Measure-Object).count)`n`n" -ForegroundColor Green
    $splitRawContent=$rawContent -split "><"
    $count=1
    foreach ($line in $splitRawContent){
        if ($dostop){break}
        if ($line -match "a href=`"https://.*$searchCriteria.*-product.*</a"){
            
            #Parseing description from line
            $desc=$line -split "View Details`">"
            $desc=$desc[1]
            $desc=$desc -split "</a"

            # Finding "Add to cart" text under the listing, searching UNTIL criteria specified.
            $count2=1
            $OOS=$true
            do{
                if ($splitRawContent[($count+$count2)] -match "Add to cart"){
                    #There is an add to cart button!
                    $OOS=$false
                    break
                }
                $count2++
            } until ($splitRawContent[($count+$count2)] -match "item-compare-box")

            if (-not($OOS)){
                
                # Ignoring entries
                $doContinue=$false
                foreach ($ignore in $searchIgnores){
                    if ($desc -like $ignore) { 
                        if ($showIgnores){write-host Ignoring: $desc}
                        $doContinue=$true 
                    }
                }
                if ($doContinue){continue}


                write-host `n(date)
                write-host The following item has `"Add to cart`" under it:`n$desc`n -ForegroundColor green

                try{
                    #Getting product ID from search page rawcontent/html
                    $count2=0
                    do{
                        if ($splitRawContent[($count+$count2)] -match "\>Item \#\:"){
                            $productID=$splitRawContent[($count+$count2)]
                            $productID=$productID -split "\>Item \#\: \<\/strong\>"
                            $productID=$productID[-1]
                            $productID=$productID -split "</li"
                            $productID=$productID[0]
                            write-host Product ID: $productID
                            #Item is out of stock
                            break
                        }
                        $count2++
                    } until ($splitRawContent[($count+$count2)] -match "item-compare-box")
                            

                    # Stop the current loop if in testing mode.
                    if ($testing){continue}

                    # noting items found for $stopInStckAll logic below
                    $itemsFound=$true


                    write-host "  Attempting to add to cart"
                    start-process $chromeLocation "https://secure.$neweggDomain/Shopping/AddtoCart.aspx?Submit=ADD&ItemList=$productID"

                    if ($playAlertOnStock){
                        # stopping in case multiple cart adds happen.  No reason to have the alert sound trying to play multiple times at once.
                        stop-alert
                        write-host "  Playing alert sound"
                        $soundJob=invoke-alert -times $playAlertTimes -wav "C:\windows\media\Alarm02.wav"
                        write-host "  --You can stop the alert sound by typing `"stop-alert`"."
                    }

                    if ($doTelegram){
                        write-host "  Sending alert via Telegram"
                        $Message = "Newegg Bot tried to add the following item to your cart: $productID, $desc"
                        $Response = Invoke-RestMethod -Uri "https://api.telegram.org/bot$($TG_BOT_TOKEN)/sendMessage?chat_id=$($TG_BOT_CHAT_ID )&text=$($Message)"
                        #$Message = "$url"
                        #$Response = Invoke-RestMethod -Uri "https://api.telegram.org/bot$($TG_BOT_TOKEN)/sendMessage?chat_id=$($TG_BOT_CHAT_ID )&text=$($Message)"
                    }

                    # If set to $true, don't continue to search and add other items to cart once it's been attempted.  
                    if ($stopInStockOne){
                        write-host "`nScript is configured to stop after stock found.  Stopping."
                        $doStop=$true
                        break
                    } #if stopinstockOne
                    
                }catch{
                    Write-Warning "Something went wrong trying to parse: $url"
                    $pageContent
                    $_
                }


            } else {
                Write-host "No stock found on search results for: $desc`n"
            } # if line notmatch outofstock
        
            
         }
        $count++
    } # foreach line in splitrawcontent
    # If set to $true, don't look for new items in stock.  
    if ($stopInStockAll -and (-not($stopInStockOne)) -and $itemsFound){
        write-host "`nScript is configured to stop after all current items in stock have tried ot add to cart.  Stopping."
        $doStop=$true
    } #if stopinstockOne
    if(-not($doStop)){
        Write-host "`nWaiting configured delay between checks: $delay sec"
        sleep $delay 
    }   
} #while


read-host "`n`nPress enter to stop alert(if needed) and finish script."
stop-alert

