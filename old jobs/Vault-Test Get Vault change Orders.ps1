########### Script to get active change orders #########


import-module powervault 


open-vaultconnection -server "vault-test" -vault "vault-test" -user "nolle" -password "Password"

#uses a list someone will ahve to manage to get active changeorders, Working on a way to get all this data automatically

Get-Content "C:\Vault Data\Changeorders.txt" | ForEach-Object {$vault.ChangeOrderService.GetChangeOrderByNumber($_)} | Export-Csv "C:\Vault Data\Change Order List.csv"
