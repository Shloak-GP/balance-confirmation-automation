# Letting the owner try it through ngrok (temporary link)

1. Put a password in `local_overrides.json` next to run_real.py (copy local_overrides.example.json):
       {"access_password": "pick-a-long-password-here"}
   Without it the link has NO login and anyone who gets the URL can see customer balances.
2. Start everything with one command:  python run_ngrok.py   (it also creates the password and prints the link)
3. In a second window:  ngrok http 8080   (ngrok must be signed in once: ngrok config add-authtoken ...).
4. Send him the https://....ngrok-free.app address and the password (separately, e.g. by phone/WhatsApp).
   ngrok may show a "Visit site" warning page the first time: he clicks it.
5. Stop it when done: Ctrl+C in both windows. The link dies with ngrok.

Know before you send it:
- Everything he sees and types goes through ngrok's servers. Test with the sample data, or get your boss's OK first.
- Your computer must stay on and awake. Saves land in YOUR copy of the Excel, not his.
- "Open Excel" and the Outlook buttons do nothing for him (they act on your computer). He can use Download.
- It is a trial: one tester at a time, and a new link every time ngrok restarts (free plan).
