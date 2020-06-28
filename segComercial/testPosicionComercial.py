import requests

url = "https://eu1.unwiredlabs.com/v2/process.php"


mmc = raw_input("MMC: ")
mnc = raw_input("MNC: ")
lac = raw_input("LAC: ")
cellId = raw_input("Cell ID: ")

payload = "{\"token\": \"xxxxx\",\"radio\": \"gsm\",\"mcc\": {MMC},\"mnc\": {MNC},\"cells\": [{\"lac\": {LAC},\"cid\": {CELLID}}],\"address\": 1}"
cadPayload = payload.replace("{MMC}", mmc)
cadPayload = cadPayload.replace("{MNC}", mnc)
cadPayload = cadPayload.replace("{LAC}", lac)
cadPayload = cadPayload.replace("{CELLID}", cellId)
response = requests.request("POST", url, data=cadPayload)

print(response.text)