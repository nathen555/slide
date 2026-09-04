# SlideMiner

A browser-only Google Slides extractor for assignment research. Paste a Google Slides URL, authenticate with a Google API key or OAuth access token, or download the deck from your school account and upload the `.pptx` file. Export the result as a study guide in Markdown.

## Run it

Open [index.html](index.html) directly in a browser. No build step or server is required.

## Connect Google Slides

1. In Google Cloud Console, create or select a project.
2. Enable the **Google Slides API**.
3. Create an API key and restrict it to the Slides API and the domains where you will host this page.
4. Paste the key into **Connection settings**.

For private presentations, use an OAuth 2.0 access token with permission to read the presentation. The page keeps credentials in memory only and sends requests directly to Google. Do not commit keys or tokens to the repository.

If your school account cannot use Google Cloud, download the presentation while signed into your school account: **File → Download → Microsoft PowerPoint (.pptx)**. Then use the upload box in SlideMiner. PowerPoint files are parsed locally in your browser.

You can also upload PDFs or scanned `JPG`, `PNG`, and `WEBP` images. SlideMiner renders PDF pages and runs English OCR in the browser, then sends the recognized text through the same quick-answer and study-guide views. OCR may need an internet connection to load the browser libraries, but the document itself is processed locally.

The file picker accepts multiple files at once, including mixed PowerPoints, PDFs, and images. SlideMiner compiles them into one study set while labeling each slide or page with its source filename. You can also paste multiple public Google Slides URLs into the URL box, one per line; linked decks are compiled together and labeled by source.

For stronger context-aware summaries, add a Gemini API key under **Connection settings**. The optional AI pass reads the entire extracted deck and returns only supported main points, quiz-worthy terms, and named people with roles. Without a key, the local context analysis is used and no deck text is sent to Gemini.

The deck must be shared with the Google account represented by the key/token. This tool does not bypass Google permissions.
