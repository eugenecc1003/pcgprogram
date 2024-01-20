from pcgfunction.function_zo13 import *
import time
import os
from flask import *

import sys
sys.path.append("..")

function_zo13 = Blueprint('function_zo13', __name__)


@function_zo13.route("/sample_ZO13", methods=["GET", "POST"])
def sample_ZO13():
    if 'zo13' not in session["OK"][3]:
        return render_template('pcc.html',
                               alert="You don't have the authorization. Please contact Eugene!!",
                               nickname=session["OK"][0],
                               email=session["OK"][1])
    else:
        return redirect(url_for("sampledownload", outputfilename="sample_ZO13_rawdata.xlsx"))


@function_zo13.route("/zo13vlookup", methods=["GET", "POST"])
def zo13vlookup():
    if 'zo13' not in session["OK"][3]:
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
        zo13rawdatafile = request.files.getlist("zo13rawdata")[0]
        cudatafile = request.files.getlist("cudata")[0]
        formatcheck = ('.xlsx', '.xls', '.XLSX', '.XLS')

        try:
            zo13rawdatafileformat = zo13rawdatafile.filename[zo13rawdatafile.filename.index(
                '.'):]
            cudatafileformat = cudatafile.filename[cudatafile.filename.index(
                '.'):]

            if (zo13rawdatafileformat.lower() in formatcheck) & (cudatafileformat.lower() in formatcheck):
                format = '.xlsx'
            else:
                return render_template('pcc.html',
                                       alert='format is not xlsx or xls !!',
                                       nickname=session["OK"][0],
                                       email=session["OK"][1])
            zo13rawdatafile.save(userfoldertimepath+"/zo13rawdata"+format)
            cudatafile.save(userfoldertimepath+"/cudata"+format)
        except:
            return render_template('pcc.html',
                                   alert="Please choose the file!!",
                                   nickname=session["OK"][0],
                                   email=session["OK"][1])
        outputfilename = zo13_vlookup(userfoldertimepath)
        return redirect(url_for("download",
                                userfoldertime=userfoldertime,
                                outputfilename=outputfilename))


@function_zo13.route("/zo13transfer", methods=["GET", "POST"])
def zo13transfer():
    if 'zo13' not in session["OK"][3]:
        return render_template('pcc.html',
                               alert="You don't have the authorization. Please contact Eugene!!",
                               nickname=session["OK"][0],
                               email=session["OK"][1])
    else:
        org = request.form["organization"]
        channel = request.form["channel"]
        fty = request.form["factory"]
        basepath = "./public/uploads"
        userfolderpath = os.path.join(basepath, session["OK"][2])
        userfoldertime = str(time.time()).replace('.', '')
        userfoldertimepath = os.path.join(userfolderpath, userfoldertime)
        os.mkdir(userfoldertimepath)
        zo13transferfile = request.files.getlist("zo13transfer")[0]
        formatcheck = ('.xlsx', '.xls', '.XLSX', '.XLS')

        try:
            zo13transferfileformat = zo13transferfile.filename[zo13transferfile.filename.index(
                '.'):]

            if zo13transferfileformat.lower() in formatcheck:
                format = '.xlsx'
            else:
                return render_template('pcc.html',
                                       alert='format is not xlsx or xls !!',
                                       nickname=session["OK"][0],
                                       email=session["OK"][1])
            zo13transferfile.save(userfoldertimepath+"/zo13transfer"+format)
        except:
            return render_template('pcc.html',
                                   alert="Please choose the file!!",
                                   nickname=session["OK"][0],
                                   email=session["OK"][1])
        outputfilename = ZO13_transfer(userfoldertimepath, org, channel, fty)
        return redirect(url_for("download",
                                userfoldertime=userfoldertime,
                                outputfilename=outputfilename))
