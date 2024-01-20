from linebot import (LineBotApi, WebhookHandler)
from linebot.exceptions import (InvalidSignatureError)
from linebot.models import *

from secretstoken import *

# Channel Access Token
line_bot_api = LineBotApi(CHANNEL_ACCESS_TOKEN)  # 你的Channel AcessToken
# Channel Secret
handler = WebhookHandler(CHANNEL_SECRET)  # 你的Channel Secret
