#!../../bin/darwin-aarch64/test01

#- SPDX-FileCopyrightText: 2005 Argonne National Laboratory
#-
#- SPDX-License-Identifier: EPICS

#- You may have to change test01 to something else
#- everywhere it appears in this file

#< envPaths

epicsEnvSet ("STREAM_PROTOCOL_PATH", "../../protocols")

## Register all support components
dbLoadDatabase "../../dbd/test01.dbd"
test01_registerRecordDeviceDriver(pdbbase)

## Load record instances
dbLoadRecords("../../db/test01.db", "SYS=val")
dbLoadRecords("../../db/test02.db")

cd "$ASYN"

cd "$TOP"

iocInit()

## Start any sequence programs
#seq snctest01,"user=1h7"
