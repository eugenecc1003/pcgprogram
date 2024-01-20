from pcgfunction.function_608A import *
import time
import os
from flask import *

import sys
sys.path.append("..")

function_608A = Blueprint('function_608A', __name__)


@function_608A.route("/sample_29_barcode", methods=["GET", "POST"])
def sample_29_barcode():
    if ('6080' not in session["OK"][3]) & (('608A' not in session["OK"][3])):
        return render_template('pcc.html',
                               alert="You don't have the authorization. Please contact Eugene!!",
                               nickname=session["OK"][0],
                               email=session["OK"][1])
    else:
        return redirect(url_for("sampledownload", outputfilename="sample_29_BarcodeFromAPL.xlsx"))


@function_608A.route("/sample_29_ZRSD1002M18", methods=["GET", "POST"])
def sample_29_ZRSD1002M18():
    if ('6080' not in session["OK"][3]) & (('608A' not in session["OK"][3])):
        return render_template('pcc.html',
                               alert="You don't have the authorization. Please contact Eugene!!",
                               nickname=session["OK"][0],
                               email=session["OK"][1])
    else:
        return redirect(url_for("sampledownload", outputfilename="sample_29_NWGW_ZRSD1002M18.xlsx"))


@function_608A.route("/sample_29_ZRSD1002P2", methods=["GET", "POST"])
def sample_29_ZRSD1002P2():
    if ('6080' not in session["OK"][3]) & (('608A' not in session["OK"][3])):
        return render_template('pcc.html',
                               alert="You don't have the authorization. Please contact Eugene!!",
                               nickname=session["OK"][0],
                               email=session["OK"][1])
    else:
        return redirect(url_for("sampledownload", outputfilename="sample_29_NWGW_ZRSD1002P2.xls"))


@function_608A.route("/barcode29", methods=["GET", "POST"])
def barcode29():
    if ('6080' not in session["OK"][3]) & (('608A' not in session["OK"][3])):
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
        barcode29 = request.files.getlist("barcode29")[0]
        try:
            barcode29format = barcode29.filename[barcode29.filename.index(
                '.'):]
            if (barcode29format.lower() in ('.xlsx', '.xls')):
                format = '.xlsx'
            else:
                return render_template('pcc.html', alert='format is not xlsx or xls !!', nickname=session["OK"][0], email=session["OK"][1])
            barcode29.save(userfoldertimepath+"/BarcodeFromAPL_29"+format)
        except:
            return render_template('pcc.html', alert="Please choose the file!!", nickname=session["OK"][0], email=session["OK"][1])
        outputfilename = barcode_29(userfoldertimepath)
        return redirect(url_for("download",
                                userfoldertime=userfoldertime,
                                outputfilename=outputfilename))


@function_608A.route("/NWGW29", methods=["GET", "POST"])
def NWGW29():
    if ('6080' not in session["OK"][3]) & (('608A' not in session["OK"][3])):
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

        tolerance = request.form["tolerance"]
        file_M18_29 = request.files.getlist("M18_29")[0]

        try:
            M18_29_format = file_M18_29.filename[file_M18_29.filename.index(
                '.'):]

            if (M18_29_format.lower() in ('.xlsx', '.xls')):
                format = '.xlsx'
            else:
                return render_template('pcc.html', alert='format is not xlsx or xls !!', nickname=session["OK"][0], email=session["OK"][1])
            file_M18_29.save(userfoldertimepath+"/M18_29"+format)

        except:
            return render_template('pcc.html', alert="Please choose the file!!", nickname=session["OK"][0], email=session["OK"][1])

        flist = request.files.getlist("P2_29")
        for f in flist:
            try:
                # 格式是從點往後的值
                format = f.filename[f.filename.index('.'):]
                if format.lower() in ('.xlsx', '.xls'):
                    format = '.xlsx'
                else:
                    # 出現問題
                    return render_template('pcc.html', alert='format is wrong!!', nickname=session["OK"][0], email=session["OK"][1])

                f.save(userfoldertimepath+"/P2_29_"+str(flist.index(f))+format)

            except:
                # 出現問題
                return render_template('pcc.html', alert="Please choose the file!!", nickname=session["OK"][0], email=session["OK"][1])

        outputfilename = NWGW_29(userfoldertimepath, tolerance)
        return redirect(url_for("download",
                                userfoldertime=userfoldertime,
                                outputfilename=outputfilename))
