from flask import Flask

app = Flask(__name__)
app.secret_key = 'kgsfintechconnectkey'

import KGSweb.views

if __name__ == "__main__":
	app.run()