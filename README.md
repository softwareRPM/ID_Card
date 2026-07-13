# Digital Contact Card

## Main QR-code URL

Generate the QR code for the main contact page:

`https://softwareRPM.github.io/ID_Card/`

This page shows the three buttons:

- **Call Me** — opens the phone dialer.
- **Email Me** — opens the default email app.
- **Save Contact** — downloads or opens `contact.vcf`.

Update the placeholder telephone number and email address in `index.html` and `contact.vcf` before publishing.

## Publish with GitHub Pages

1. Commit and push the project to the `master` branch.
2. In the GitHub repository, open **Settings → Pages**.
3. Under **Build and deployment**, choose **GitHub Actions** as the source.
4. Pushes to `master` will publish the site automatically. The first deployment usually takes a few minutes.

After the deployment succeeds, use the main URL above in the QR code. It will continue to work whenever you update and republish the page.

## Optional updatable PDF

The permanent PDF QR address is `https://softwareRPM.github.io/ID_Card/pdf/`.

Add your PDF as `document.pdf`. Replace that exact file whenever it changes; the PDF QR code stays the same.
