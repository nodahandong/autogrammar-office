# AutoGrammar for Outlook, Word and Excel

One package adds AutoGrammar to **Outlook, Word and Excel** on Windows, Mac and the web.

- **Outlook:** while **writing** an email, use **Fix grammar** (one click, fixed in place) or the **AutoGrammar** panel. While **reading** an email, click **Translate** to read it in English, Korean, Khmer or another language.
- **Word:** **Home → AutoGrammar** opens a panel. Select text, pick a style, then click **Replace**.
- **Excel:** select cells with text, pick a style, then click **Apply to cells**. Formulas and numbers are left alone, and there's an **Undo** button.

Every panel has the same choices: Fix grammar, Business, Natural, Friendly, Shorter, Clearer, To English, To Korean, To Khmer, plus **Other languages** (Japanese, Chinese, Spanish, Portuguese). To just read a translation, pick a language and use **Copy** (or simply read it) instead of Replace.

> **Using Outlook or Office in Chrome?** The Chrome extension already works on outlook.office.com. For Word/Excel on the web, the add-in below is better.

Office add-ins have to be loaded from a website, so setup has two parts: put the files on a free GitHub Pages site (once), then add the add-in to each app. If you connect GitHub in this project, I can do Part 1 for you.

---

## Part 1: Put the files online (once, about 10 minutes)

The add-in is set up for **https://nodahandong.github.io/autogrammar-office/**.

1. Sign in at **https://github.com** as **nodahandong**.
2. Click **+** (top-right) → **New repository**. Name it exactly **autogrammar-office**, choose **Public**, tick **Add a README file**, and click **Create repository**.
3. Unzip **AutoGrammar-Office.zip**. In the repository, click **Add file → Upload files** and drag in **everything inside** the `office-addin` folder (including the `assets` folder). Click **Commit changes**.
4. Go to **Settings → Pages**. Under *Branch*, choose **main** and **/(root)**, then click **Save**.
5. After 1–2 minutes, open **https://nodahandong.github.io/autogrammar-office/office.html**. If you see the AutoGrammar page, it worked.

## Part 2a: Outlook (Windows, Mac, web)

1. Open **https://aka.ms/olksideload** and sign in with your Outlook account.
2. Go to **My add-ins → Custom add-ins → Add a custom add-in → Add from file** and choose **manifest-outlook.xml**.
3. Close Outlook completely and open it again (the desktop app can take a few minutes to pick up a new add-in).
   - **Writing:** click **New mail**. **Fix grammar** and **AutoGrammar** are on the message ribbon. In **new Outlook** and on the web, click the **Apps** icon (four squares) in the message toolbar.
   - **Reading:** open an email. **Translate** is on the ribbon (new Outlook: **Apps** icon, or the **…** menu at the top of the email).
4. Click **AutoGrammar**, open **Settings**, paste your OpenAI API key, and click **Save**.

## Part 2b: Word and Excel

The file to use is **manifest-word-excel.xml**. One file covers both apps.

**On the web (office.com):** open a document → **Home → Add-ins → More Add-ins** (or **Insert → Add-ins**) → **My Add-ins** → **Upload My Add-in** → choose the file.

**On Mac:**
1. In Finder, press **Cmd + Shift + G** and go to `~/Library/Containers/com.microsoft.Word/Data/Documents/`
2. Create a folder named **wef** there (if it isn't there already) and copy **manifest-word-excel.xml** into it.
3. Do the same for Excel: `~/Library/Containers/com.microsoft.Excel/Data/Documents/wef`
4. Restart Word/Excel. Go to **Insert → Add-ins → My Add-ins** (or **Home → Add-ins**) and choose **AutoGrammar** from the developer add-ins.

**On Windows (easy way):** in the unzipped `office-addin` folder, double-click **Install-Word-Excel.bat** and click **Yes** when Windows asks for permission. Then follow the 4 steps it shows. If it can't share the folder (some company PCs block this), use the manual steps below or Word/Excel on office.com.

**On Windows (manual):**
1. Make a folder, for example `C:\AutoGrammar`, and copy **manifest-word-excel.xml** into it.
2. Right-click the folder → **Properties → Sharing → Share…** → **Share**. Note the network path it shows (like `\\YOUR-PC\AutoGrammar`).
3. In Word: **File → Options → Trust Center → Trust Center Settings → Trusted Add-in Catalogs**. Paste the network path into *Catalog Url*, click **Add catalog**, tick **Show in Menu**, click **OK**, and restart Word.
4. **Insert → My Add-ins → SHARED FOLDER → AutoGrammar → Add**. Do the same in Excel (the trust setting is shared, so you can usually skip to this step).

Then click **AutoGrammar** on the **Home** tab, open **Settings**, paste your API key, and click **Save** (once per app).

## Outlook: I installed it but don't see it

1. Check you're in the right place: the writing buttons only appear in a **new mail / reply** window, and **Translate** only when an email is open.
2. Open **https://nodahandong.github.io/autogrammar-office/taskpane.html** in your browser. If it says *404*, Part 1 isn't finished (upload all files and turn on Pages). Outlook can't show the add-in without it.
3. Close Outlook fully and reopen it. Still nothing? Go to **aka.ms/olksideload → My add-ins**, remove AutoGrammar, add **manifest-outlook.xml** again, and restart Outlook.
4. **Updating:** when you get a new version, upload the new files to GitHub (they replace the old ones), then remove and re-add **manifest-outlook.xml**.

## Notes

- Word: the new text keeps line breaks, but bold or colors inside the selected part become plain text. Select whole sentences.
- Excel: at most 200 text cells per run. Each cell is checked separately.
- Work or school accounts: if you can't add custom add-ins, your IT admin has blocked them.
