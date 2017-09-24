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

void Hello(const Nan::FunctionCallbackInfo<v8::Value> &info) {
    double inReal[1] = {0.2};
    double outReal[1];
    TA_Integer outBeg;
    TA_Integer outNbElement;
//    TA_COS(0, 1, inReal, &outBeg, &outNbElement, outReal);
    TA_SAR(0, 1, inReal, inReal, 0, 0, &outBeg, &outNbElement, outReal);
    v8::Local<v8::Array> array = Nan::New<v8::Array>(1);
    Nan::Set(array, 0, Nan::New<v8::Number>(outReal[0]));
    info.GetReturnValue().Set(array);
}

void Init(v8::Local<v8::Object> exports, v8::Local<v8::Object> module) {
    exports->Set(Nan::New("hello").ToLocalChecked(), Nan::New<v8::FunctionTemplate>(Hello)->GetFunction());
}

NODE_MODULE(talib_binding, Init)