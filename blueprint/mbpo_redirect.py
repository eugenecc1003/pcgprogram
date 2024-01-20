from pcgfunction.mbpo import *
import time
import os
from flask import *

import sys
sys.path.append("..")

mbpoapp = Blueprint('mbpoapp', __name__)


@mbpoapp.route("/MBPO27", methods=["GET", "POST"])
def MBPO27():
    if 'MBPO27' not in session["OK"][3]:
        return render_template('pcc.html',
                               alert="You don't have the authorization. Please contact Eugene!!",
                               nickname=session["OK"][0],
                               email=session["OK"][1])
    else:
        # basepath = os.path.join(os.path.dirname(__file__), 'public', 'uploads')
        basepath = "./public/uploads"
        # 確認登入後，檢查有沒有sessionmongoid專屬資料夾，沒有則新
        userfolderpath = os.path.join(basepath, session["OK"][2])
        # if request.method == 'POST':
        # 選取檔案的清單
        flist = request.files.getlist("file[]")
        userfoldertime = str(time.time()).replace('.', '')
        # session["OK"].insert(4, userfoldertime)
        userfoldertimepath = os.path.join(userfolderpath, userfoldertime)
        os.mkdir(userfoldertimepath)
        for f in flist:
            try:
                # 格式是從點往後的值
                format = f.filename[f.filename.index('.'):]
                filetime = str(time.time()).replace('.', '')
                if format.lower() in ('.xlsx', '.xls'):
                    format = '.xlsx'

                else:
                    # 出現問題
                    return render_template('pcc.html', alert='format is wrong!!', nickname=session["OK"][0], email=session["OK"][1])

                f.save(userfoldertimepath+"/"+filetime+format)

            except:
                # 出現問題
                return render_template('pcc.html', alert="Please choose the file!!", nickname=session["OK"][0], email=session["OK"][1])

        outputfilename = MBPO_27(userfoldertimepath)

        # global filepath
        # filepath = "./public/uploads/" + sessionmongoid+"/"+userfoldertime

        # 刪除檔案
        # @after_this_request
        # def removeFile(response):
        #    for timepath in os.scandir(userfolderpath):
        #        for file in os.scandir(timepath.path):
        #            os.remove(file.path)
        #        os.rmdir(timepath.path)
        #    return response

        return redirect(url_for("download",
                                userfoldertime=userfoldertime,
                                outputfilename=outputfilename))


@mbpoapp.route("/MBPO29", methods=["GET", "POST"])
def MBPO29():
    if 'MBPO29' not in session["OK"][3]:
        return render_template('pcc.html',
                               alert="You don't have the authorization. Please contact Eugene!!",
                               nickname=session["OK"][0],
                               email=session["OK"][1])
    else:
        basepath = "./public/uploads"
        userfolderpath = os.path.join(basepath, session["OK"][2])
        userfoldertime = str(time.time()).replace('.', '')
        userfoldertimepath = os.path.join(userfolderpath, userfoldertime)
        os.mkdir(userfoldertimepath)

        option = request.form["option"]
        fileAPL29 = request.files.getlist("APL29")[0]
        filePF29 = request.files.getlist("PF29")[0]
        fileCC29 = request.files.getlist("CC29")[0]
        try:
            APL29format = fileAPL29.filename[fileAPL29.filename.index('.'):]
            PF29format = filePF29.filename[filePF29.filename.index('.'):]
            CC29format = fileCC29.filename[fileCC29.filename.index('.'):]
            if (APL29format.lower() in ('.xlsx', '.xls')) & (PF29format.lower() in ('.xlsx', '.xls')) & (CC29format.lower() in ('.xlsx', '.xls')):
                format = '.xlsx'
            else:
                return render_template('pcc.html', alert='format is not xlsx or xls !!', nickname=session["OK"][0], email=session["OK"][1])
            fileAPL29.save(userfoldertimepath+"/APL29"+format)
            filePF29.save(userfoldertimepath+"/PF29"+format)
            fileCC29.save(userfoldertimepath+"/CC29"+format)
        except:
            return render_template('pcc.html', alert="Please choose the file!!", nickname=session["OK"][0], email=session["OK"][1])

        outputfilename = MBPO_29(userfoldertimepath, option)

        return redirect(url_for("download",
                                userfoldertime=userfoldertime,
                                outputfilename=outputfilename))


@mbpoapp.route("/MBPO38", methods=["GET", "POST"])
def MBPO38():
    if 'MBPO38' not in session["OK"][3]:
        return render_template('pcc.html',
                               alert="You don't have the authorization. Please contact Eugene!!",
                               nickname=session["OK"][0],
                               email=session["OK"][1])
    else:
        basepath = "./public/uploads"
        userfolderpath = os.path.join(basepath, session["OK"][2])
        userfoldertime = str(time.time()).replace('.', '')
        userfoldertimepath = os.path.join(userfolderpath, userfoldertime)
        os.mkdir(userfoldertimepath)

        fileGTN38 = request.files.getlist("GTN38")[0]
        filePF38 = request.files.getlist("PF38")[0]
        fileCC38 = request.files.getlist("CC38")[0]
        try:
            GTN38format = fileGTN38.filename[fileGTN38.filename.index('.'):]
            PF38format = filePF38.filename[filePF38.filename.index('.'):]
            CC38format = fileCC38.filename[fileCC38.filename.index('.'):]
            if (GTN38format.lower() in ('.xlsx', '.xls')) & (PF38format.lower() in ('.xlsx', '.xls')) & (CC38format.lower() in ('.xlsx', '.xls')):
                format = '.xlsx'
            else:
                return render_template('pcc.html', alert='format is not xlsx or xls !!', nickname=session["OK"][0], email=session["OK"][1])
            fileGTN38.save(userfoldertimepath+"/GTN38"+format)
            filePF38.save(userfoldertimepath+"/PF38"+format)
            fileCC38.save(userfoldertimepath+"/CC38"+format)
        except:
            return render_template('pcc.html', alert="Please choose the file!!", nickname=session["OK"][0], email=session["OK"][1])

        outputfilename = MBPO_38(userfoldertimepath)

        return redirect(url_for("download",
                                userfoldertime=userfoldertime,
                                outputfilename=outputfilename))


# SAMPLE 27


@mbpoapp.route("/sample_27_GTN", methods=["GET", "POST"])
def sample_27_GTN():
    if 'MBPO27' not in session["OK"][3]:
        return render_template('pcc.html',
                               alert="You don't have the authorization. Please contact Eugene!!",
                               nickname=session["OK"][0],
                               email=session["OK"][1])
    else:
        return redirect(url_for("sampledownload", outputfilename="sample_27_GTN.xlsx"))

# SAMPLE 29


@mbpoapp.route("/sample_29_APL", methods=["GET", "POST"])
def sample_29_APL():
    if 'MBPO29' not in session["OK"][3]:
        return render_template('pcc.html',
                               alert="You don't have the authorization. Please contact Eugene!!",
                               nickname=session["OK"][0],
                               email=session["OK"][1])
    else:
        return redirect(url_for("sampledownload", outputfilename="sample_29_APL.xlsx"))


@mbpoapp.route("/sample_29_MBPO_CustomerCode", methods=["GET", "POST"])
def sample_29_CustomerCode():
    if 'MBPO29' not in session["OK"][3]:
        return render_template('pcc.html',
                               alert="You don't have the authorization. Please contact Eugene!!",
                               nickname=session["OK"][0],
                               email=session["OK"][1])
    else:
        return redirect(url_for("sampledownload", outputfilename="sample_29_MBPO_CustomerCode.xlsx"))


@mbpoapp.route("/sample_29_MBPO_PartnerFunction", methods=["GET", "POST"])
def sample_29_PartnerFunction():
    if 'MBPO29' not in session["OK"][3]:
        return render_template('pcc.html',
                               alert="You don't have the authorization. Please contact Eugene!!",
                               nickname=session["OK"][0],
                               email=session["OK"][1])
    else:
        return redirect(url_for("sampledownload", outputfilename="sample_29_MBPO_PartnerFunction.xlsx"))

# SAMPLE 38


@mbpoapp.route("/sample_38_GTN", methods=["GET", "POST"])
def sample_38_GTN():
    if 'MBPO38' not in session["OK"][3]:
        return render_template('pcc.html',
                               alert="You don't have the authorization. Please contact Eugene!!",
                               nickname=session["OK"][0],
                               email=session["OK"][1])
    else:
        return redirect(url_for("sampledownload", outputfilename="sample_38_GTN.xlsx"))


@mbpoapp.route("/sample_38_MBPO_CustomerCode", methods=["GET", "POST"])
def sample_38_CustomerCode():
    if 'MBPO38' not in session["OK"][3]:
        return render_template('pcc.html',
                               alert="You don't have the authorization. Please contact Eugene!!",
                               nickname=session["OK"][0],
                               email=session["OK"][1])
    else:
        return redirect(url_for("sampledownload", outputfilename="sample_38_MBPO_CustomerCode.xlsx"))


@mbpoapp.route("/sample_38_MBPO_PartnerFunction", methods=["GET", "POST"])
def sample_38_PartnerFunction():
    if 'MBPO38' not in session["OK"][3]:
        return render_template('pcc.html',
                               alert="You don't have the authorization. Please contact Eugene!!",
                               nickname=session["OK"][0],
                               email=session["OK"][1])
    else:
        return redirect(url_for("sampledownload", outputfilename="sample_38_MBPO_PartnerFunction.xlsx"))
