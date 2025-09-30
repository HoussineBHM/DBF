from pyngrok import ngrok, conf
import time

logfile = '/workspaces/DBF/ngrok.log'

# Start an http tunnel on port 8501
http_tunnel = ngrok.connect(8501, "http")
url = http_tunnel.public_url
with open(logfile, 'w') as f:
    f.write(f'Public URL: {url}\n')
    f.write(str(http_tunnel))

print(f'Public URL: {url}')
# keep process alive to maintain tunnel
try:
    while True:
        time.sleep(60)
except KeyboardInterrupt:
    ngrok.disconnect(http_tunnel.public_url)
    ngrok.kill()
