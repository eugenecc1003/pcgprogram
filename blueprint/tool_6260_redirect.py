from pcgfunction.function_6260 import *
import time
import os
from flask import *

import sys
sys.path.append("..")

function_6260 = Blueprint('function_6260', __name__)


@function_6260.route("/sample_6260_ZRSD1039", methods=["GET", "POST"])
def sample_6260_ZRSD1039():
    if '6260' not in session["OK"][3]:
        return render_template('pcc.html',
                               alert="You don't have the authorization. Please contact Eugene!!",
                               nickname=session["OK"][0],
                               email=session["OK"][1])
    else:
        return redirect(url_for("sampledownload", outputfilename="sample_6260_ZRSD1039.XLSX"))


@function_6260.route("/sample_6260_ZRMM0002", methods=["GET", "POST"])
def sample_6260_ZRMM0002():
    if '6260' not in session["OK"][3]:
        return render_template('pcc.html',
                               alert="You don't have the authorization. Please contact Eugene!!",
                               nickname=session["OK"][0],
                               email=session["OK"][1])
    else:
        return redirect(url_for("sampledownload", outputfilename="sample_6260_ZRMM0002.XLSX"))


@function_6260.route("/dnprint6260", methods=["GET", "POST"])
def dnprint6260():
    if '6260' not in session["OK"][3]:
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
        zrsd1039file = request.files.getlist("zrsd1039")[0]
        zrmm0002file = request.files.getlist("zrmm0002")[0]
        formatcheck = ('.xlsx', '.xls', '.XLSX', '.XLS')
        try:
            zrsd1039fileformat = zrsd1039file.filename[zrsd1039file.filename.index(
                '.'):]
            zrmm0002fileformat = zrmm0002file.filename[zrmm0002file.filename.index(
                '.'):]
            if (zrsd1039fileformat.lower() in formatcheck) & (zrmm0002fileformat.lower() in formatcheck):
                format = '.xlsx'
            else:
                return render_template('pcc.html',
                                       alert='format is not xlsx or xls !!',
                                       nickname=session["OK"][0],
                                       email=session["OK"][1])
            zrsd1039file.save(userfoldertimepath+"/zrsd1039"+format)
            zrmm0002file.save(userfoldertimepath+"/zrmm0002"+format)
        except:
            return render_template('pcc.html',
                                   alert="Please choose the file!!",
                                   nickname=session["OK"][0],
                                   email=session["OK"][1])
        outputfilename = DN_print_6260(userfoldertimepath)
        return redirect(url_for("download",
                                userfoldertime=userfoldertime,
                                outputfilename=outputfilename))
