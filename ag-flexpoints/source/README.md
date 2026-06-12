# source/ — manual fallback drop-zone

If the automated portal login fails, save the FlexPoints page here from a
logged-in browser (Ctrl+S → "Webpage, HTML Only", or copy-paste the page text
into a `.txt` file). `fetch_flexpoints.py` uses the newest file in this folder
when it can't log in itself.

Everything in this folder except this README is gitignored — saved portal
pages never get committed.
