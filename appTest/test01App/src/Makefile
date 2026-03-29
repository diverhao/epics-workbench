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
DBD += test01.dbd

test01_DBD += base.dbd
test01_DBD += modules.dbd

test01_SRCS += test01_registerRecordDeviceDriver.cpp
test01_SRCS += devAdd10.c
test01_SRCS += registrarHello.c

test01_SRCS_DEFAULT += test01Main.cpp
test01_SRCS_vxWorks += -nil-

test01_LIBS += $(EPICS_BASE_IOC_LIBS)

include $(TOP)/configure/RULES

