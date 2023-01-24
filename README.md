A python script that scrapes Gmail for .xls file convert to .xlsx and upload to one drive.

Configuration:
    time[int][default = 5]: the amount of time to wait in hours before scrapping again.
    clear_cache[bool][default=false]: clear recorded data of previously scrapped messages.
    logout_google[bool][default=false]: logs you out of your gmail account.

configurations are made in config.json file