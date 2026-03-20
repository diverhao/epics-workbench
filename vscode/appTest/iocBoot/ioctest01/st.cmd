#!../../bin/darwin-aarch64/test01

#- SPDX-FileCopyrightText: 2005 Argonne National Laboratory
#-
#- SPDX-License-Identifier: EPICS

#- You may have to change test01 to something else
#- everywhere it appears in this file

# < envPaths

epicsEnvSet("STREAM_PROTOCOL_PATH", "$STREAM_DEVICE/streamApp")

## Register all support components
dbLoadDatabase "../../dbd/test01.dbd"
test01_registerRecordDeviceDriver(pdbbase) 

#dbLoadRecords("../../db/test01.db", "P=$(P=CEA:), R=$(P=LAB:)")
dbLoadRecords("../../test01App/Db/4g.db", "S=,BUS=")
dbLoadRecords("../../db/test03.db")
cd "$ASYN1"
dbLoadRecords("db/devInt32.db", "P=$(P=CEA:), R=$(P=LAB:)")

iocInit()

dbpf("asyndevAiInt32A0", "22")
dbpf("abcd1", "33")

## Start any sequence programs
#seq snctest01,"user=1h7"
