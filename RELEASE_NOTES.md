## PDF Mule v1.7.0-rc2 — release candidate

### English
This is a **test version** (release candidate) of PDF Mule 1.7.0.
- **Fixed:** Translate PDF stopped with "You made too many requests to the server" on longer documents. PDF Mule now sends the text to Google calmly, one part at a time, and when Google asks to slow down it waits and tries again by itself.
- If Google still refuses for a long time, the translated parts are saved (…_PARTIAL.txt) and the next run **continues where it stopped**.
- Everything from rc1: programs install themselves (Tesseract, Ghostscript, LibreOffice), PDF to Word fix for printed books, HTML to PDF with Edge/Chrome, iPhone photos, clearer errors, step-by-step tool windows.

### Ελληνικά
Αυτή είναι **δοκιμαστική έκδοση** (release candidate) του PDF Mule 1.7.0.
- **Διόρθωση:** η Μετάφραση PDF σταματούσε με «You made too many requests to the server» σε μεγάλα έγγραφα. Τώρα το PDF Mule στέλνει το κείμενο στο Google ήρεμα, ένα κομμάτι τη φορά, και όταν το Google ζητά να πάει πιο αργά, περιμένει και ξαναδοκιμάζει μόνο του.
- Αν το Google αρνείται για πολλή ώρα, τα μεταφρασμένα κομμάτια αποθηκεύονται (…_PARTIAL.txt) και η επόμενη εκτέλεση **συνεχίζει από εκεί που σταμάτησε**.
- Όλα όσα είχε η rc1: τα προγράμματα εγκαθίστανται μόνα τους (Tesseract, Ghostscript, LibreOffice), διόρθωση PDF σε Word για τυπωμένα βιβλία, HTML σε PDF με Edge/Chrome, φωτογραφίες iPhone, κατανοητά μηνύματα λάθους, παράθυρα εργαλείων σε βήματα.
