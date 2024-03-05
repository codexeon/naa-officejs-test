## Setup

    npm install
    npx office-addin-dev-certs install

Create a .ENV file with the following:
```
SSL_CRT_FILE=C:\Users\<username>\.office-addin-dev-certs\localhost.crt
SSL_KEY_FILE=C:\Users\<username>\.office-addin-dev-certs\localhost.key
PORT=443
HTTPS=true
```

Set the following registry key to sideload manifest:
```
[HKEY_CURRENT_USER\Software\Microsoft\Office\16.0\WEF\Developer]
"manifest1"=<path to manifest on disk>
```

## Start the dev server
    npm start