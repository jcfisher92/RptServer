#!/bin/sh

RPTSERVER=/usr/local/rptserver
LIB=$RPTSERVER/lib/*

#
# Set the axis2 lib  directory
AXIS2_LIB=/opt/axis2/lib/*

#
# Set the poi library
POI_LIB=/opt/poi/*:/opt/poi/lib/*:/opt/poi/ooxml-lib/*

#
# local dir
CLASS_PATH=$LIB:$AXIS2_LIB:$RPTSERVER:$POI_LIB

/usr/bin/java -cp $CLASS_PATH:com.emerywaterhouse.rpt.server.RptServer.class -server -Dsqlite.purejava=true -Xms100m -Xmx3000m -XX:MaxPermSize=1024m -XX:+CMSClassUnloadingEnabled -XX:+UseConcMarkSweepGC com.emerywaterhouse.rpt.server.RptServer
