# PPTX Remote

Present a PowerPoint deck from any computer and control it from your phone.

**👉 https://aidenjeong8.github.io/pptx-remote/**

![Presenter page](docs/img/presenter.png)

## How to use it

1. Open the site on the computer you're presenting from and drop in a `.pptx` file.
2. Scan the QR code with your phone's camera.
3. Click **Present**. Your phone now has back/next buttons, your speaker notes, and a list of
   slides by title that you can tap to jump to any one.

There's nothing to install and no account. The computer and phone don't need to be on the same
network.

## How it works

Everything runs in the browser. The site is a static GitHub Pages site with no backend.

- **Rendering:** the `.pptx` is parsed and drawn in the browser by
  [`@aiden0z/pptx-renderer`](https://github.com/aiden0z/pptx-renderer). Speaker notes are read
  from the file with JSZip.
- **Phone link:** the presenter page and the phone connect directly with WebRTC through
  [PeerJS](https://peerjs.com). PeerJS's free public server only introduces the two devices.
  The deck then goes straight from the computer to the phone over an encrypted data channel,
  and the phone renders its own copy for thumbnails and notes. Slide changes travel the same way,
  in both directions.

```
docs/
  index.html    landing page + presenter
  remote.html   phone remote
  js/deck.js    shared: load a .pptx, render slides, read notes
  vendor/       pinned copies of pptx-renderer 1.3.0, PeerJS 1.5.5, JSZip 3.10.1, qrcode-generator 1.4.4
```

To run it locally, serve `docs/` with any static server, for example `cd docs && python3 -m http.server`.

## Limitations

- Animations and transitions aren't played: each slide shows in its finished state.
- Fonts that aren't embedded in the file or installed on the computer are replaced with similar ones.
- Browsers only allow fullscreen after a click on that computer, so start the presentation there.
- On some strict networks (certain cellular carriers or corporate firewalls) a direct WebRTC
  connection can't be made. Putting both devices on the same Wi-Fi fixes it.

## License

MIT. Bundled third-party libraries keep their own licenses; see `docs/vendor/`.
