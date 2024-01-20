import os
from flask import *
from dotenv import load_dotenv
from flask_cors import CORS

from connectfunction.function_linebot import *
from connectfunction.function_mongo import *

from blueprint.mbpo_redirect import mbpoapp
from blueprint.tool_608A_redirect import function_608A
from blueprint.tool_6260_redirect import function_6260
from blueprint.tool_zo13_redirect import function_zo13

from secretstoken import *
# Load environment variables from .env file
load_dotenv()

app = Flask(__name__, static_folder='public', static_url_path='/')
app.secret_key = FLASK_SECRET_KEY  # 設定session要用
CORS(app)  # Enable CORS for all routes and origins

# Configure FLASK_DEBUG from environment variable
app.config['DEBUG'] = os.environ.get('FLASK_DEBUG')

app.register_blueprint(mbpoapp, url_prefix='')
app.register_blueprint(function_608A, url_prefix='')
app.register_blueprint(function_6260, url_prefix='')
app.register_blueprint(function_zo13, url_prefix='')


@app.route('/')
def newindex():
    return render_template('newindex.html')


@app.route('/error')  # /error?msg=錯誤訊息
def error():
    message = request.args.get(
        'msg', 'something wrong?? Please contact Eugene!!')
    return render_template('error.html', messageoutput=message)


@app.route("/signup", methods=["POST"])
def signup():
    # 從前端接收資料
    nickname = request.form['nickname']
    email = request.form['email']
    password = request.form['password']
    authorization = []
    # 根據接收到的資料收到資料庫
    # 檢查email
    # collection = db.user
    result = collection.find_one({"email": email})
    if email == "":
        return redirect("/error?msg=the email cannot be block.")
    elif result != None:
        return redirect("/error?msg=the email has been used already.")
    # 把資料放進資料庫
    collection.insert_one({
        'status': 'check',
        'nickname': nickname,
        'email': email,
        'password': password,
        'authorization': authorization
    })
    return redirect("/")


@app.route("/signin", methods=["POST"])
def signin():
    email = request.form['email']
    password = request.form['password']

    # collection = db.user
    # 檢查密碼
    result = collection.find_one({
        "$and": [
            {"email": email},
            # {"password": password}
        ]
    })
    if result == None:
        return redirect("/error?msg=You haven't sign up this email!!")
    else:
        if result['status'] == 'check':
            return redirect("/error?msg=authorization is checking. Please contact Eugene!!")
        elif result['password'] != password:
            return redirect("/error?msg=password is wrong!!")

    session["OK"] = [result['nickname'], result['email'],
                     str(result['_id']), result['authorization']]
    return redirect("/pcc")


@app.route("/signout")
def signout():
    sessionmongoid = session["OK"][2]
    basepath = os.path.join(os.path.dirname(__file__), 'public', 'uploads')
    userfolderpath = os.path.join(basepath, sessionmongoid)
    for timepath in os.scandir(userfolderpath):
        for file in os.scandir(timepath.path):
            os.remove(file.path)
        os.rmdir(timepath.path)

    del session["OK"]
    return redirect("/")


@app.route('/pcc', methods=['GET', 'POST'])
def pcc():
    # 如果沒有登入，直接導回主畫面
    if "OK" in session:  # by每一次連線確認，只要有，就當有登入
        sessionnickname = session["OK"][0]
        sessionemail = session["OK"][1]
        sessionmongoid = session["OK"][2]
        # sessionauthorization = session["OK"][3]
        # os.path.dirname(path) 去掉檔名，返回目錄
        # __file__表示了當前檔案的path
        # originpath = os.path.dirname(__file__)
        basepath = os.path.join(os.path.dirname(__file__), 'public', 'uploads')
        # 確認登入後，檢查有沒有sessionmongoid專屬資料夾，沒有則新增
        global userfolderpath
        userfolderpath = os.path.join(basepath, sessionmongoid)
        # session["OK"].insert(3, userfolderpath)
        if not os.path.isdir(userfolderpath):
            os.mkdir(userfolderpath)

            # return send_from_directory(userfoldertimepath, outputfilename, as_attachment=True)

        return render_template('pcc.html',  nickname=sessionnickname, email=sessionemail)
    # 沒有登入的話，回首頁
    else:
        return redirect("/")


@app.route("/download/<userfoldertime>/<outputfilename>")
def download(userfoldertime, outputfilename):
    filepath = "./public/uploads/" + session["OK"][2]+"/"+userfoldertime
    return send_from_directory(filepath, outputfilename, as_attachment=True)


@app.route("/sampledownload/<outputfilename>")
def sampledownload(outputfilename):
    filepath = "./public/sample"
    return send_from_directory(filepath, outputfilename, as_attachment=True)


# 監聽所有來自 /callback 的 Post Request
@app.route("/callback", methods=['POST'])
def callback():
    # get X-Line-Signature header value
    signature = request.headers['X-Line-Signature']
    # get request body as text
    body = request.get_data(as_text=True)
    print(body)

    app.logger.info("Request body: " + body)
    # handle webhook body
    try:
        handler.handle(body, signature)
    except InvalidSignatureError:
        abort(400)
    return 'OK'


# 處理訊息
@handler.add(MessageEvent, message=TextMessage)
def handle_message(event):
    # event.message.text 就是用戶傳來的文字訊息
    msg = event.message.text.split(' ')
    # 文字回應
    if msg[0] in ['CHECK', 'Check', 'check']:
        message = TextSendMessage(text=check())

    elif (msg[0] in emaillist()) & (msg[1] in ['APPROVE', 'Approve', 'approve']):
        message = TextSendMessage(text=approve(msg[0]))

    elif (msg[0] in emaillist()) & (msg[1] in ['INSERT', 'Insert', 'insert']):
        message = TextSendMessage(text=authorization_insert(msg[0], msg[2]))

    else:
        message = TextSendMessage(text="????")
    line_bot_api.reply_message(event.reply_token, message)


@handler.add(PostbackEvent)
def handle_message(event):
    print(event.postback.data)


if __name__ == "__main__":  # 如果以主程式執行
    app.run(port=9527)  # 立刻執行伺服器
