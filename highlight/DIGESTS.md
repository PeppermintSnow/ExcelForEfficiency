## Subresource Integrity

If you are loading Highlight.js via CDN you may wish to use [Subresource Integrity](https://developer.mozilla.org/en-US/docs/Web/Security/Subresource_Integrity) to guarantee that you are using a legimitate build of the library.

To do this you simply need to add the `integrity` attribute for each JavaScript file you download via CDN. These digests are used by the browser to confirm the files downloaded have not been modified.

```html
<script
  src="//cdnjs.cloudflare.com/ajax/libs/highlight.js/11.9.0/highlight.min.js"
  integrity="sha384-9mu2JKpUImscOMmwjm1y6MA2YsW3amSoFNYwKeUHxaXYKQ1naywWmamEGMdviEen"></script>
<!-- including any other grammars you might need to load -->
<script
  src="//cdnjs.cloudflare.com/ajax/libs/highlight.js/11.9.0/languages/go.min.js"
  integrity="sha384-WmGkHEmwSI19EhTfO1nrSk3RziUQKRWg3vO0Ur3VYZjWvJRdRnX4/scQg+S2w1fI"></script>
```

The full list of digests for every file can be found below.

### Digests

```
sha384-B6XNN+tQFI6cb/bizGIoHqASsFpNSzhB4rXEhX6IrjEdHLeJaR6Onmx2+ArhvnCm /es/languages/vbnet.js
sha384-LdqDRVWxc7cPE2ISZcBIeHoQAyfnXmEArz0layp8fOjceievL8eX9cV+HgneOkQ4 /es/languages/vbnet.min.js
sha384-NRGkScJ2K7VJUsHL7H9eHncQMh99wPeeOYyPkHZiUfihAg+Xb0j1JotJI3DkqX5f /languages/vbnet.js
sha384-+y0KLxbRrWxqnfGRhWWQTHHEwDd1OhkOacgf5QfJa+5ydoCf3SWObb+XrzxBSfqa /languages/vbnet.min.js
sha384-0jvkRPYWT2l0cM1vEStI24kzUSquJFGbcDq1eIMXJKRYEIfKuKfWFNVZxrdau+iD /highlight.js
sha384-8gb4ilGfJpMd8/wpjooPUOmKO6nBemsKtmxUXZUmY7r6+6zG5wtK2VHrM5IcqOm2 /highlight.min.js
```

