import os, shutil, pathlib, argparse

import pytz
import datetime
from win32com.propsys import propsys, pscon
from PIL import Image

mediaEndings = "png,jpg,jpeg,tiff,bmp,gif,264,3g2,3gp,3gp2,3gpp,3gpp2,3mm,3p2,60d,787,89,aaf,aec,aep,aepx,aet,aetx,ajp,ale,am,amc,amv,amx,anim,aqt,arcut,arf,asf,asx,avb,avc,avd,avi,avp,avs,avs,avv,axm,bdm,bdmv,bdt2,bdt3,bik,bin,bix,bmk,bnp,box,bs4,bsf,bvr,byu,camproj,camrec,camv,ced,cel,cine,cip,clpi,cmmp,cmmtpl,cmproj,cmrec,cpi,cst,cvc,cx3,d2v,d3v,dat,dav,dce,dck,dcr,dcr,ddat,dif,dir,divx,dlx,dmb,dmsd,dmsd3d,dmsm,dmsm3d,dmss,dmx,dnc,dpa,dpg,dream,dsy,dv,dv-avi,dv4,dvdmedia,dvr,dvr-ms,dvx,dxr,dzm,dzp,dzt,edl,evo,eye,ezt,f4p,f4v,fbr,fbr,fbz,fcp,fcproject,ffd,flc,flh,fli,flv,flx,gfp,gl,gom,grasp,gts,gvi,gvp,h264,hdmov,hkm,ifo,imovieproj,imovieproject,ircp,irf,ism,ismc,ismv,iva,ivf,ivr,ivs,izz,izzy,jss,jts,jtv,k3g,kmv,ktn,lrec,lsf,lsx,m15,m1pg,m1v,m21,m21,m2a,m2p,m2t,m2ts,m2v,m4e,m4u,m4v,m75,mani,meta,mgv,mj2,mjp,mjpg,mk3d,mkv,mmv,mnv,mob,mod,modd,moff,moi,moov,mov,movie,mp21,mp21,mp2v,mp4,mp4v,mpe,mpeg,mpeg1,mpeg4,mpf,mpg,mpg2,mpgindex,mpl,mpl,mpls,mpsub,mpv,mpv2,mqv,msdvd,mse,msh,mswmm,mts,mtv,mvb,mvc,mvd,mve,mvex,mvp,mvp,mvy,mxf,mxv,mys,ncor,nsv,nut,nuv,nvc,ogm,ogv,ogx,osp,otrkey,pac,par,pds,pgi,photoshow,piv,pjs,playlist,plproj,pmf,pmv,pns,ppj,prel,pro,prproj,prtl,psb,psh,pssd,pva,pvr,pxv,qt,qtch,qtindex,qtl,qtm,qtz,r3d,rcd,rcproject,rdb,rec,rm,rmd,rmd,rmp,rms,rmv,rmvb,roq,rp,rsx,rts,rts,rum,rv,rvid,rvl,sbk,sbt,scc,scm,scm,scn,screenflow,sec,sedprj,seq,sfd,sfvidcap,siv,smi,smi,smil,smk,sml,smv,spl,sqz,srt,ssf,ssm,stl,str,stx,svi,swf,swi,swt,tda3mt,tdx,thp,tivo,tix,tod,tp,tp0,tpd,tpr,trp,ts,tsp,ttxt,tvs,usf,usm,vc1,vcpf,vcr,vcv,vdo,vdr,vdx,veg,vem,vep,vf,vft,vfw,vfz,vgz,vid,video,viewlet,viv,vivo,vlab,vob,vp3,vp6,vp7,vpj,vro,vs4,vse,vsp,w32,wcp,webm,wlmp,wm,wmd,wmmp,wmv,wmx,wot,wp3,wpl,wtv,wve,wvx,xej,xel,xesc,xfl,xlmv,xmv,xvid,y4m,yog,yuv,zeg,zm1,zm2,zm3,zmv"

parser = argparse.ArgumentParser(formatter_class=argparse.ArgumentDefaultsHelpFormatter)
parser.add_argument(
    "-i", "--InputPath", help="Path to start at", type=pathlib.Path, required=True
)
parser.add_argument("-o", "--OutputPath", help="Path to output at", type=pathlib.Path)
parser.add_argument(
    "--no-recursion",
    dest="Recursion",
    help="Should Script Run through sub Direcory's",
    action=argparse.BooleanOptionalAction,
    default=True,
)
parser.add_argument(
    "--FileTypes",
    help="The File Type/Types you want to Extract. Deliminate multiple with ','",
    default=mediaEndings,
)
parser.add_argument(
    "--Standardize",
    help="Should Script Standardize File Names?",
    action=argparse.BooleanOptionalAction,
    default=False,
)
config = vars(parser.parse_args())

desiredEndings = tuple(config["FileTypes"].split(","))
if config["OutputPath"] != None:
    outputPath = config["OutputPath"]
else:
    outputPath = os.path.join(os.path.dirname(config["InputPath"]), "Extracted")


def getDesiredType(paths):
    return [path for path in paths if path.lower().endswith(desiredEndings)]


def getMedia(path, media=None):
    if media == None:
        media = []
    os.chdir(path)
    currentDirList = os.listdir()
    desiredDir = getDesiredType(currentDirList)
    desiredDir = [
        pathlib.Path(os.path.join(path, mediaName)) for mediaName in desiredDir
    ]
    media.extend(desiredDir)
    if config["Recursion"]:
        nextDirs = [i for i in currentDirList if os.path.isdir(i)]
        for nextDirName in nextDirs:
            nextDir = os.path.join(path, nextDirName)
            getMedia(nextDir, media)
    return media


def getMediaCreationDate(path):
    dt = None
    try:
        exif = Image.open(path).getexif()
        if exif != None:
            date = exif.get(306)
            if date != None:
                dt = date
    except:
        try:
            properties = propsys.SHGetPropertyStoreFromParsingName(str(path))
            dt = properties.GetValue(pscon.PKEY_Media_DateEncoded).GetValue()
            dt = dt.astimezone(pytz.FixedOffset(120))
            dt = str(dt)
        except:
            return None
    if dt != None:
        return "{0}{1}{2}_{3}{4}{5}".format(
            dt[0:4], dt[5:7], dt[8:10], dt[11:13], dt[14:16], dt[17:19]
        )
    else:
        return None


def makeUniquePath(path):
    i = 1
    tmpPath = path
    while os.path.exists(tmpPath):
        tmpPath = path
        nameEnd = tmpPath.find(".")
        tmpPath = tmpPath[:nameEnd] + "_{}".format(i) + tmpPath[nameEnd:]
        i += 1
    return tmpPath


media = getMedia(config["InputPath"])
if not os.path.exists(outputPath):
    os.makedirs(outputPath)

i = 1
for m in media:
    name = m.name
    if config["Standardize"]:
        name = getMediaCreationDate(m)
        if name == None:
            name = "NODATE"
        extensionBegining = m.name.find(".")
        name += m.name[extensionBegining:]

    dest_dir = os.path.join(outputPath, name)
    dest_dir = makeUniquePath(dest_dir)
    shutil.copy(m, dest_dir)
    print("{}/{} copying {} -> {}".format(i, len(media), m.name, dest_dir))
    i += 1

print("...DONE...")
