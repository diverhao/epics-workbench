#include <stdio.h>
#include <epicsExport.h>
#include <iocsh.h>

void hello() {
    printf("Welcome to EPICS Workbench!\n");
}

static const iocshArg    *helloArgs[] = {};
static const iocshFuncDef helloFuncDef = {
    "hello",
    0,
    helloArgs,
};

static void helloCallFunc(const iocshArgBuf *args) {
    hello();
}

static void helloRegistrar(void) {
    iocshRegister(&helloFuncDef, helloCallFunc);
}
epicsExportRegistrar(helloRegistrar);
