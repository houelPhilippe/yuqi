import webview
import threading
import time
import sys

html_content = """
<!DOCTYPE html>
<html>
<head>
  <script src="turndown.min.js"></script>
</head>
<body>
  <div id="test1">------ texte</div>
  <div id="test2">- - - - - - - texte</div>
  <script>
    window.onload = function() {
      try {
        const t = new TurndownService({ bulletListMarker: '-' });
        const res1 = t.turndown(document.getElementById('test1').innerHTML);
        const res2 = t.turndown(document.getElementById('test2').innerHTML);
        window.pywebview.api.log(res1, res2);
      } catch (e) {
        window.pywebview.api.log("Error", e.toString());
      }
    };
  </script>
</body>
</html>
"""

class Api:
    def log(self, res1, res2):
        print("RES1:", repr(res1))
        print("RES2:", repr(res2))
        with open("test_out.txt", "w") as f:
            f.write(res1 + "\\n" + res2)
        window.destroy()

if __name__ == '__main__':
    api = Api()
    window = webview.create_window('Test', html=html_content, js_api=api)
    webview.start()
