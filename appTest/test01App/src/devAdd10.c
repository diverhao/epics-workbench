
#include <stdlib.h>
#include <stddef.h>
#include <stdio.h>
#include <string.h>

#include <dbCommon.h>
#include <aiRecord.h>
#include <devSup.h>
#include <epicsExport.h>
#include <dbAccess.h>
#include <dbStaticLib.h>

static long report_device(int);
static long init_device(int);
static long init_record(dbCommon *);
static long read_ai(aiRecord *);
static long add_record(struct dbCommon *prec);
static long delete_record(struct dbCommon *prec);
static void get_record_field_addr(char *channelName, char *fieldName, aiRecord *prec);

static struct
{
    /**
     * For details see devSup.h `typedef struct typed_dset {...}`
     *
     * when we choose USE_TYPED_DSET, `dset` will choose the
     * typed verison, otherwise will choose untyped version
     * dset =
     * {
     *     long number;
     *     long (*report)(int level);
     *     long (*init)(int pass); // called twice in iocInit()
     *     long (*init_record)(struct dbCommon *prec); // init record
     *     long (*get_ioint_info)(int detach, struct dbCommon *prec, IOSCANPVT *pscan); // rarely used
     * }
     */
    dset common;                     // we choose USED_TYPE_DSET
    long (*read_ai)(aiRecord *prec); // specifically for this record type
    long (*special_linconv)(aiRecord *rec);
} DevAdd10 = {
    {
        6,
        report_device,
        init_device,
        init_record,
        NULL,
    },
    read_ai,
    NULL};

/**
 * this the **device support extension table**
 *
 * It contains 2 methods, the first one is invoked after init_device() and before init_record()
 * in doResolveLinks() in initDatabase() in dbLockInitRecords() in iocBuild_2() in iocInit.c
 *
 * etherIP uses it to allocate `prec->dpvt` space, then parse the INP/SCAN/... and fill in
 * the `prec->dpvt`
 */
static struct dsxt extended_table = {add_record, delete_record};

long add_record(struct dbCommon *prec)
{
    printf("Link record %s with device type DevAdd10\n\n", prec->name);
    return 0;
}

long delete_record(struct dbCommon *prec)
{
    printf("Un-link record %s with device type DevAdd10\n\n", prec->name);
    return 0;
}

/**
 * Called in dbior()
 */
static long report_device(int level)
{
    printf("This is an example device support Add10\n\n");
    return 1;
}

/**
 * called twice in iocInit(), before and after the init_record()
 *
 * for each device support, no matter how many records, only called twice
 *
 * One is at initDevSup(), the next is at finishDevSup()
 *
 * In between it is record, database, link startup
 */
static long init_device(int pass)
{
    printf("init deivce Add10: pass %d\n\n", pass);
    if (pass == 0)
    {
        // we **usually must** register extension table here
        devExtend(&extended_table);
    }

    // initialize the device, you can do whatever you want
    // ...
    return 0;
}

/**
 * called in iocInit(), called for each record
 */
static long init_record(dbCommon *prec)
{
    aiRecord *pai = (aiRecord *)prec;
    printf("init device Add10 record %s \n\n", prec->name);

    // 2 ways to read the INP string in db file
    // (1) static access
    // Read the INP field text using dbStaticLib
    DBENTRY dbEntry;
    dbInitEntry(pdbbase, &dbEntry);
    if (dbFindRecord(&dbEntry, prec->name) == 0)
    {
        if (dbFindField(&dbEntry, "INP") == 0)
        {
            char *linkText = dbGetString(&dbEntry);
            if (linkText != NULL)
            {
                // "@ABC DEF 33"
                printf("INP link text: %s\n\n", linkText);
            }
        }
    }
    dbFinishEntry(&dbEntry);
    // (2) runtime access
    DBLINK a = pai->inp;
    // "ABC DEF 33", the "@" is removed becuse it is considered as instio
    printf("INP link text: %s\n\n", a.value.instio.string);

    // Make sure record processing routine does not perform any conversion
    pai->linr = 0;
    return (0);
}

/**
 * called everything this record is processed
 */
static long read_ai(aiRecord *prec)
{
    aiRecord *pai = (aiRecord *)prec;
    printf("Process device Add10 record %s\n\n", prec->name);
    if (pai->inp.text != NULL)
    {
        printf("INP link text: %s\n\n", pai->inp.text);
    }
    // driver support can be invoked here
    pai->val = pai->val + 10;
    pai->udf = FALSE;

    // get_record_field_addr("AAA", "INP", prec);

    return (2);
}

epicsExportAddress(dset, DevAdd10);


