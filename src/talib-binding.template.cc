#include <nan.h>
#include "../ta-lib/c/include/ta_libc.h"

/*!
 *
 * Copyright 2017 - acrazing
 *
 * @author acrazing joking.young@gmail.com
 * @since 2017-09-24 23:05:39
 * @version 1.0.0
 * @desc talib-binding.generated.cc
 */

static const v8::Local<v8::String> inOpenName = Nan::New("Open").ToLocalChecked();
static const v8::Local<v8::String> inHighName = Nan::New("High").ToLocalChecked();
static const v8::Local<v8::String> inLowName = Nan::New("Low").ToLocalChecked();
static const v8::Local<v8::String> inCloseName = Nan::New("Close").ToLocalChecked();
static const v8::Local<v8::String> inVolumeName = Nan::New("Volume").ToLocalChecked();

void TA_FUNC_SAR(const Nan::FunctionCallbackInfo<v8::Value> &info) {
    int usedArgc = 0;
    v8::Local<v8::Array> inFirst = v8::Local<v8::Array>::Cast(info[usedArgc++]);
    int firstIsRecord = 0;
    if (inFirst->Length() > 0 && inFirst->Get(0)->IsObject()) {
        firstIsRecord = 1;
    }
    v8::Local<v8::Array> inHigh_JS, inLow_JS;
    if (firstIsRecord == 0) {
        inHigh_JS = inFirst;
        inLow_JS = v8::Local<v8::Array>::Cast(info[usedArgc++]);
    }
    uint32_t i = 0;
    uint32_t inLength = inFirst->Length();
    double optInAcceleration = TA_REAL_DEFAULT;
    double optInMaximum = TA_REAL_DEFAULT;
    int startIdx = 0;
    // The end index is included
    int endIdx = inLength - 1;
    int outBegIdx = 0;
    int outNBElement = 0;
    int argc = info.Length();
    if (argc > usedArgc && info[usedArgc]->IsInt32()) {
        optInAcceleration = info[usedArgc]->NumberValue();
    }
    usedArgc++;
    if (argc > usedArgc && info[usedArgc]->IsNumber()) {
        optInMaximum = info[usedArgc++]->NumberValue();
    }
    usedArgc++;
    if (argc > usedArgc && info[usedArgc]->IsInt32()) {
        startIdx = info[usedArgc++]->Int32Value();
    }
    info[0]->ToString();
    usedArgc++;
    if (argc > usedArgc && info[usedArgc]->IsInt32()) {
        endIdx = info[usedArgc]->Int32Value();
        if (endIdx >= (int) inLength) {
            Nan::ThrowRangeError("The end index should be less than input data length(max is length - 1)");
            return;
        }
    }
    int lookback = TA_SAR_Lookback(optInAcceleration, optInMaximum);
    int temp = lookback > startIdx ? lookback : startIdx;
    int allocationSize = temp > endIdx ? 0 : endIdx - temp + 1;
    if (allocationSize == 0) {
        info.GetReturnValue().Set(Nan::New<v8::Array>(0));
        return;
    }
    double *inHigh = new double[inLength];
    double *inLow = new double[inLength];
    double *outReal = new double[allocationSize];
    if (firstIsRecord) {
        v8::Local<v8::Object> record;
        for (i = 0; i < inLength; i++) {
            record = inFirst->Get(i)->ToObject();
            inHigh[i] = record->Get(inHighName)->NumberValue();
            inLow[i] = record->Get(inLowName)->NumberValue();
        }
    } else {
        for (i = 0; i < inLength; i++) {
            inHigh[i] = inHigh_JS->Get(i)->NumberValue();
            inLow[i] = inLow_JS->Get(i)->NumberValue();
        }
    }
    TA_RetCode result = TA_SAR(startIdx, endIdx, inHigh, inLow, optInAcceleration, optInMaximum, &outBegIdx,
                               &outNBElement, outReal);
    if (result != TA_SUCCESS) {
        delete[] inHigh;
        delete[] inLow;
        delete[] outReal;
        TA_RetCodeInfo retCodeInfo;
        TA_SetRetCodeInfo(result, &retCodeInfo);
        char error[100];
        strcat(error, "TA_SAR ERROR: ");
        strcat(error, retCodeInfo.enumStr);
        strcat(error, " - ");
        strcat(error, retCodeInfo.infoStr);
        Nan::ThrowError(error);
        return;
    }
    v8::Local<v8::Array> outReal_JS = Nan::New<v8::Array>(allocationSize);
    for (i = 0; i < (uint32_t) allocationSize; i++) {
        outReal_JS->Set(i, Nan::New<v8::Number>(outReal[i]));
    }
    info.GetReturnValue().Set(outReal_JS);
    delete[] inHigh;
    delete[] inLow;
    delete[] outReal;
}

void Init(v8::Local<v8::Object> exports, v8::Local<v8::Object> module) {
    exports->Set(Nan::New("SAR").ToLocalChecked(), Nan::New<v8::FunctionTemplate>(TA_FUNC_SAR)->GetFunction());
}

NODE_MODULE(talib_binding, Init)