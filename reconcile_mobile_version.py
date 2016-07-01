#Script to reconcile and post MOBILE_EDIT_VERSION 
import arcpy, os, datetime, sys
from arcpy import env
arcpy.env.overwriteOutput = True

def RecPost():
    #set workspace, copied this from Carls code and not sure why we need sdeadmin as workspace
    sde = "Database Connections/sdeadmin.sde"
    RPUDwkspace = "Database Connections/RPUD.sde"
    arcpy.env.workspace = sde

    # get time for naming log file
    ReconcileTime = datetime.datetime.now().strftime("%Y-%m-%d-%H-%M")
    filePath = "" # somewhere to store the log file

    #reconcile, post, detect conflict by attribute(column)
    arcpy.ReconcileVersions_management(RPUDwkspace, "ALL_VERSIONS", "SDE.DEFAULT", "RPUD.MOBILE_EDIT_VERSION", "LOCK_ACQUIRED", "NO_ABORT", "BY_ATTRIBUTE", "FAVOR_TARGET_VERSION", "POST")
    print "Rec&Post MOBILE_VERSION is Completed"

RecPost()