## PDF Mule v1.6.1 — the programs it needs install themselves

### English
- **No more "Tesseract OCR was not found":** the first time you use OCR, PDF Mule installs Tesseract by itself, with English, Greek and page orientation. Other OCR languages download by themselves when a document needs them.
- **Ghostscript** (PDF/A, strong compression) and **LibreOffice** (Word / PowerPoint / Excel → PDF) also install by themselves the first time a tool needs them, with progress in the tool's log.
- **Windows:** one "Yes" at the Windows permission prompt. **Mac:** Homebrew is used; without it, Terminal opens and installs it (type your Mac password once). **Linux:** your password in the system dialog.
- **Help → Dependency status → Install missing programs** installs all of them at once.
- Every download is checked (Windows signature, Mac code signature) before it is installed.

### Ελληνικά
- **Τέλος το «Tesseract OCR was not found»:** την πρώτη φορά που χρησιμοποιείς OCR, το PDF Mule εγκαθιστά μόνο του το Tesseract, με Αγγλικά, Ελληνικά και αναγνώριση προσανατολισμού σελίδας. Οι άλλες γλώσσες OCR κατεβαίνουν μόνες τους όταν τις χρειαστεί ένα έγγραφο.
- Το **Ghostscript** (PDF/A, δυνατή συμπίεση) και το **LibreOffice** (Word / PowerPoint / Excel → PDF) εγκαθίστανται επίσης μόνα τους την πρώτη φορά που τα χρειάζεται ένα εργαλείο, με την πρόοδο στο log του εργαλείου.
- **Windows:** ένα «Ναι» στο παράθυρο αδειών των Windows. **Mac:** χρησιμοποιείται το Homebrew· αν δεν υπάρχει, ανοίγει το Terminal και το εγκαθιστά (γράφεις μία φορά τον κωδικό του Mac). **Linux:** ο κωδικός σου στο παράθυρο του συστήματος.
- **Help → Dependency status → Install missing programs** τα εγκαθιστά όλα μαζί.
- Κάθε λήψη ελέγχεται (υπογραφή Windows, υπογραφή κώδικα Mac) πριν εγκατασταθεί.
