# Postkasselevering
Generates ICS file for when the norwegian postal service will deliver to your mailbox.

# Cron
```
# Run at 2 AM and place the ics file in /var/www
0 2 * * * /home/powershell/pwsh /home/user/postkasselevering/postkasselevering.ps1 -postnr 7010 -outfile "/var/www/post.ics"
```
