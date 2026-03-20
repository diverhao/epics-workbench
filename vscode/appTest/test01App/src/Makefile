# SPDX-FileCopyrightText: 2003 Argonne National Laboratory
#
# SPDX-License-Identifier: EPICS

TOP=../..

include $(TOP)/configure/CONFIG
#----------------------------------------
#  ADD MACRO DEFINITIONS AFTER THIS LINE
#=============================

#=============================
# Build the IOC application

PROD_IOC = test01
# test01.dbd will be created and installed
DBD += test01.dbd

# test01.dbd will be made up from these files:
test01_DBD += base.dbd
test01_DBD += menuAlarmStat.dbd waveformRecord.dbd seqRecord.dbd
test01_DBD += stream-base.dbd

# Include dbd files from all support applications:
#test01_DBD += xxx.dbd

# Add all the support libraries needed by this IOC
test01_LIBS += qsrv
test01_LIBS += Com

# test01_registerRecordDeviceDriver.cpp derives from test01.dbd
test01_SRCS += test01_registerRecordDeviceDriver.cpp

# Build the main IOC entry point on workstation OSs.
test01_SRCS_DEFAULT += test01Main.cpp
test01_SRCS_vxWorks += -nil-

# Add support from base/src/vxWorks if needed
#test01_OBJS_vxWorks += $(EPICS_BASE_BIN)/vxComLibrary

# Finally link to the EPICS Base libraries
test01_LIBS += $(EPICS_BASE_IOC_LIBS)

#===========================

include $(TOP)/configure/RULES
#----------------------------------------
#  ADD RULES AFTER THIS LINE

