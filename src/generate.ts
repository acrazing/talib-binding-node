/*!
 *
 * Copyright 2017 - acrazing
 *
 * @author acrazing joking.young@gmail.com
 * @since 2017-09-24 23:05:12
 * @version 1.0.0
 * @desc generate.ts
 */

import {autorun} from "piclick"
import * as fs from "fs"
import {parseString} from "xml2js"
import {OptionalInputArgument, OutputArgument, RequiredInputArgument, TaFuncApiXml} from "./ta_func_api.generated"
import {createMap, IMap} from "known-types"
import G = require("glob")
import {Arguments} from "yargs"

interface Json2ts {
    convertObjectToTsInterfaces<T>(content: T, name?: string): string;
    toUpperFirstLetter(text: string): string;
    toLowerFirstLetter(text: string): string;
    removeMajority(text: string): string;
}

const json2ts: Json2ts = require('json2ts')
json2ts.toLowerFirstLetter = json2ts.toUpperFirstLetter = json2ts.removeMajority = (text) => text

export function parseXml() {
    const config = fs.readFileSync(__dirname + '/../ta-lib/ta_func_api.xml', 'utf8')
    parseString(config, (err, result) => {
        if (err) {
            console.error(err)
        } else {
            fs.writeFileSync(__dirname + '/ta_func_api.generated.json', JSON.stringify(result, void 0, 2))
            const content = json2ts.convertObjectToTsInterfaces(result, 'TaFuncApiXml')
            fs.writeFileSync(__dirname + '/ta_func_api.generated.ts', content)
        }
    })
}

export function displayTypes() {
    const config: TaFuncApiXml = require('./ta_func_api.generated.json')
    const requiredNames = createMap<true>()
    const requiredTypes = createMap<true>()
    const optionalNames = createMap<true>()
    const optionalTypes = createMap<true>()
    const outputNames = createMap<true>()
    const outputTypes = createMap<true>()
    const outputFlags = createMap<true>()
    config.FinancialFunctions.FinancialFunction.forEach((func) => {
        func.RequiredInputArguments && func.RequiredInputArguments.forEach((requires) => {
            requires.RequiredInputArgument.forEach((argv) => {
                argv.Name.forEach((name) => requiredNames[name] = true)
                argv.Type.forEach((type) => requiredTypes[type] = true)
            })
        })
        func.OptionalInputArguments && func.OptionalInputArguments.forEach((options) => {
            options.OptionalInputArgument.forEach((argv) => {
                argv.Name.forEach((name) => optionalNames[name] = true)
                argv.Type.forEach((type) => optionalTypes[type] = true)
            })
        })
        func.OutputArguments && func.OutputArguments.forEach((outputs) => {
            outputs.OutputArgument.forEach((argv) => {
                argv.Name.forEach((name) => outputNames[name] = true)
                argv.Type.forEach((type) => outputTypes[type] = true)
                argv.Flags.forEach((flags) => {
                    flags.Flag.forEach((flag) => outputFlags[flag] = true)
                })
            })
        })
    })
    const result = {requiredNames, requiredTypes, optionalNames, optionalTypes, outputNames, outputTypes, outputFlags}
    process.stdout.write(Object.keys(result).map((name: keyof typeof result) => {
        return name + ': ' + Object.keys(result[name]).join(', ')
    }).join('\n'))
}

class Indenter {
    constructor(public prefix = '    ', public indented = 0) {
    }

    indent(content: string) {
        return this.prefix.repeat(++this.indented) + content
    }

    undent(content: string) {
        return this.prefix.repeat(--this.indented) + content
    }

    normal(content: string) {
        return this.prefix.repeat(this.indented) + content
    }
}

class ContentBuilder {
    constructor(public content: string[] = [], public indenter = new Indenter()) {
    }

    indent(...contents: string[]): this {
        this.indenter.indented++
        for (let i = 0; i < contents.length; i++) {
            this.content.push(this.indenter.normal(contents[i]))
        }
        return this
    }

    undent(...contents: string[]): this {
        this.indenter.indented--
        for (let i = 0; i < contents.length; i++) {
            this.content.push(this.indenter.normal(contents[i]))
        }
        return this
    }

    normal(...contents: string[]) {
        for (let i = 0; i < contents.length; i++) {
            this.content.push(this.indenter.normal(contents[i]))
        }
        return this
    }

    build() {
        return this.content.join('\n')
    }
}

export function ucFirst(str: string) {
    return str.charAt(0).toUpperCase() + str.substr(1)
}

export function normalName(name: string) {
    return name.replace(/-/g, '').replace(/\s+/g, '_')
}

export function optName(argv: OptionalInputArgument) {
    if (argv.Name.length !== 1) {
        throw new Error('out name is not a single array ' + JSON.stringify(argv))
    }
    const name = argv.Name[0]
    return normalName(name.startsWith('opt') ? name : `opt${ucFirst(name)}`)
}

export function inName(argv: RequiredInputArgument) {
    if (argv.Name.length !== 1) {
        throw new Error('out name is not a single array ' + JSON.stringify(argv))
    }
    const name = argv.Name[0]
    return normalName(name.startsWith('in') ? name : `in${ucFirst(name)}`)
}

export function outName(argv: OutputArgument) {
    if (argv.Name.length !== 1) {
        throw new Error('out name is not a single array ' + JSON.stringify(argv))
    }
    const name = argv.Name[0]
    return normalName(name.startsWith('out') ? name : `out${ucFirst(name)}`)
}

export function jsInName(argv: RequiredInputArgument) {
    return inName(argv) + '_JS'
}

export function jsOutName(argv: OutputArgument) {
    return outName(argv) + '_JS'
}

export function optDft(argv: OptionalInputArgument, ts = false) {
    const type = optType(argv)
    return type === 'TA_MAType'
        ? ts ? TA_MATypes[argv.DefaultValue[0]].split('_')[2] : TA_MATypes[argv.DefaultValue[0]]
        : argv.DefaultValue[0]
}

export function optType(argv: OptionalInputArgument, ts = false) {
    if (argv.Type.length !== 1) {
        throw new TypeError('opt type must be a single string ' + JSON.stringify(argv))
    }
    if (ts) {
        switch (argv.Type[0]) {
            case 'Integer':
            case 'Double':
                return 'number'
            case 'MA Type':
                return 'MATypes'
            default:
                throw new TypeError('Unrecognized opt param type ' + JSON.stringify(argv))
        }
    }
    switch (argv.Type[0]) {
        case 'Integer':
            return 'int'
        case 'MA Type':
            return 'TA_MAType'
        case 'Double':
            return 'double'
        default:
            throw new TypeError('Unrecognized opt param type ' + JSON.stringify(argv))
    }
}

export function outType(argv: OutputArgument) {
    if (argv.Type.length > 1) {
        throw new TypeError('output type must be a single string ' + JSON.stringify(argv))
    }
    switch (argv.Type[0]) {
        case 'Double Array':
            return 'double'
        case 'Integer Array':
            return 'int'
        default:
            throw new TypeError('unrecognized out param type ' + JSON.stringify(argv))
    }
}

export const RecordFields = ['Open', 'High', 'Low', 'Close', 'Volume']

export const TA_MATypes: IMap<string> = createMap({
    0: 'TA_MAType_SMA',
    1: 'TA_MAType_EMA',
    2: 'TA_MAType_WMA',
    3: 'TA_MAType_DEMA',
    4: 'TA_MAType_TEMA',
    5: 'TA_MAType_TRIMA',
    6: 'TA_MAType_KAMA',
    7: 'TA_MAType_MAMA',
    8: 'TA_MAType_T3',
})

export function generateBindings({_}: Arguments) {
    const body = new ContentBuilder()
    body.normal(
        '/*!',
        ' * This file is generated by generate.ts, do not edit it.',
        ' */',
        '',
        '#include <nan.h>',
        '#include "../ta-lib/c/include/ta_libc.h"',
        '',
    )
    const footer = new ContentBuilder()
    footer
        .normal('')
        .normal('void Init(v8::Local<v8::Object> exports, v8::Local<v8::Object> module) {')
        .indent('TA_RetCode retCode = TA_Initialize();')
        .normal('if (retCode != TA_SUCCESS) {')
        .indent('Nan::ThrowError("TA initialize failed!");')
        .normal('return;')
        .undent('}')
    const config: TaFuncApiXml = require('./ta_func_api.generated.json')
    config.FinancialFunctions.FinancialFunction.filter(
        (func) => _.length === 0 || _.indexOf(func.Abbreviation[0]) > -1
    ).forEach((func) => {
        const name = func.Abbreviation[0]
        const required: RequiredInputArgument[] = func.RequiredInputArguments
            ? func.RequiredInputArguments.reduce((prev, curr) => {
                return curr.RequiredInputArgument ? prev.concat(curr.RequiredInputArgument) : prev
            }, [] as RequiredInputArgument[]) : []
        const optional: OptionalInputArgument[] = func.OptionalInputArguments
            ? func.OptionalInputArguments.reduce((prev, curr) => {
                return curr.OptionalInputArgument ? prev.concat(curr.OptionalInputArgument) : prev
            }, [] as OptionalInputArgument[]) : []
        const output: OutputArgument[] = func.OutputArguments
            ? func.OutputArguments.reduce((prev, curr) => {
                return curr.OutputArgument ? prev.concat(curr.OutputArgument) : prev
            }, [] as OutputArgument[]) : []
        const doubleRequired = required.filter((argv) => RecordFields.indexOf(argv.Type[0]) === -1)
        const outInitJS: string[] = ['v8::Local<v8::Array> outAll_JS;']
        output.forEach((argv) => {
            outInitJS.push(`v8::Local<v8::Array> ${jsOutName(argv)} = Nan::New<v8::Array>(outLength);`)
        })
        if (output.length === 1) {
            outInitJS.push(`outAll_JS = ${jsOutName(output[0])};`)
        } else {
            outInitJS.push(`outAll_JS = Nan::New<v8::Array>(${output.length});`)
            output.forEach((argv, index) => {
                outInitJS.push(`outAll_JS->Set(${index}, ${jsOutName(argv)});`)
            })
        }
        const returnStatement = `info.GetReturnValue().Set(outAll_JS);`

        body
        // function declare
            .normal(`void TA_FUNC_${name}(const Nan::FunctionCallbackInfo<v8::Value> &info) {`)
            // first argument
            .indent('v8::Local<v8::Array> inFirst = v8::Local<v8::Array>::Cast(info[0]);')
            // input length
            .normal('int inLength = inFirst->Length();')
            // output length
            .normal('int outLength = 0;')
            // check input length, if is 0, return empty result, avoid check parameters
            .normal('if (inLength == 0) {')
            .indent(...outInitJS)
            .normal(returnStatement)
            .normal('return;')
            .undent('}')
            // first input is Record[] or double[]
            .normal('int firstIsRecord = 0;')
            .normal('if (inFirst->Get(0)->IsObject()) {')
            .indent('firstIsRecord = 1;')
            .undent('}')
            // declare required input
            .normal(...required.map((argv) => `double *${inName(argv)} = new double[inLength];`))
            // declare optional input
            .normal(...optional.map((argv) => {
                const type = optType(argv)
                const defaults = type === 'TA_MAType' ? TA_MATypes[argv.DefaultValue[0]] : argv.DefaultValue[0]
                return `${type} ${optName(argv)} = ${defaults};`
            }))
            // declare internal input
            .normal('int startIdx = 0;')
            .normal('int endIdx = inLength - 1;')
            .normal('int outBegIdx = 0;')
            .normal('int outNBElement = 0;')
            .normal('uint32_t i = 0;')
            .normal('int argc = info.Length();')
            // init inputs
            .normal('if (firstIsRecord == 0) {')
            .indent(...required.map((argv, index) => {
                return `v8::Local<v8::Array> ${jsInName(argv)} = v8::Local<v8::Array>::Cast(info[${index}]);`
            }))
            // init required
            .normal('for (i = 0; i < (uint32_t) inLength; i++) {')
            .indent(...required.map((argv) => `${inName(argv)}[i] = ${jsInName(argv)}->Get(i)->NumberValue();`))
            .undent('}')
            // init optional
            .normal(...optional.map((argv, index) => {
                index += required.length
                const type = optType(argv)
                const check = type === 'double' ? 'Number' : 'Int32'
                const cast = type === 'TA_MAType' ? '(TA_MAType) ' : ''
                return `${optName(argv)} = argc > ${index} && info[${index}]->Is${check}()`
                    + ` ? ${cast} info[${index}]->${check}Value() : ${optName(argv)};`
            }))
            .normal(...['startIdx', 'endIdx'].map((name, index) => {
                index += required.length + optional.length
                return `${name} = argc > ${index} && info[${index}]->IsInt32() ? info[${index}]->Int32Value() : ${name};`
            }))
            .undent('} else {')
            .indent(...required.map((argv, index) => {
                const ctor = doubleRequired.indexOf(argv) === -1
                    ? `Nan::New("${argv.Type[0]}").ToLocalChecked()`
                    : `info[${index + 1}]->ToString()`
                return `v8::Local<v8::String> ${inName(argv)}Name = ${ctor};`
            }))
            .normal(`v8::Local<v8::Object> inObject;`)
            .normal('for (i = 0; i < (uint32_t) inLength; i++) {')
            .indent('inObject = inFirst->Get(i)->ToObject();')
            .normal(...required.map((argv) => {
                return `${inName(argv)}[i] = inObject->Get(${inName(argv)}Name)->NumberValue();`
            }))
            .undent('}')
            .normal(...optional.map((argv, index) => {
                index += 1 + doubleRequired.length
                const type = optType(argv)
                const check = type === 'double' ? 'Number' : 'Int32'
                const cast = type === 'TA_MAType' ? '(TA_MAType) ' : ''
                return `${optName(argv)} = argc > ${index} && info[${index}]->Is${check}()`
                    + ` ? ${cast} info[${index}]->${check}Value() : ${optName(argv)};`
            }))
            .normal(...['startIdx', 'endIdx'].map((name, index) => {
                index += 1 + doubleRequired.length + optional.length
                return `${name} = argc > ${index} && info[${index}]->IsInt32() ? info[${index}]->Int32Value() : ${name};`
            }))
            .undent('}')
            .normal('if (startIdx < 0 || endIdx >= inLength || startIdx > endIdx) {')
            .indent('Nan::ThrowRangeError("`startIdx` or `endIdx` out of range");')
            .normal('return;')
            .undent('}')
            .normal(`int lookback = TA_${name}_Lookback(${optional.map((argv) => optName(argv)).join(', ')});`)
            .normal('int temp = lookback > startIdx ? lookback : startIdx;')
            .normal('outLength = temp > endIdx ? 0 : endIdx - temp + 1;')
            .normal('if (outLength == 0) {')
            .indent(...outInitJS)
            .normal(returnStatement)
            .normal('return;')
            .undent('}')
            .normal(...output.map((argv) => `${outType(argv)} *${outName(argv)} = new ${outType(argv)}[outLength];`))
            .normal(
                `TA_RetCode result = TA_${name}(startIdx, endIdx, `
                + required.map((argv) => inName(argv)).join(', ')
                + (required.length > 0 ? ', ' : '')
                + optional.map((argv) => optName(argv)).join(', ')
                + (optional.length > 0 ? ', ' : '')
                + '&outBegIdx, &outNBElement, '
                + output.map((argv) => outName(argv)).join(', ')
                + ');'
            )
            .normal('if (result != TA_SUCCESS) {')
            .indent(...required.map((argv) => `delete[] ${inName(argv)};`))
            .normal(...output.map((argv) => `delete[] ${outName(argv)};`))
            .normal('TA_RetCodeInfo retCodeInfo;')
            .normal('TA_SetRetCodeInfo(result, &retCodeInfo);')
            .normal('char error[100];')
            .normal(`strcpy(error, "TA_${name} ERROR: ");`)
            .normal('strcat(error, retCodeInfo.enumStr);')
            .normal('strcat(error, " - ");')
            .normal('strcat(error, retCodeInfo.infoStr);')
            .normal('Nan::ThrowError(error);')
            .normal('return;')
            .undent('}')
            .normal(...outInitJS)
            .normal('for (i = 0; i < (uint32_t) outLength; i++) {')
            .indent(...output.map((argv) => {
                return `${jsOutName(argv)}->Set(i, Nan::New<v8::Number>(${outName(argv)}[i]));`
            }))
            .undent('}')
            .normal(returnStatement)
            .normal(...required.map((argv) => `delete[] ${inName(argv)};`))
            .normal(...output.map((argv) => `delete[] ${outName(argv)};`))
            .undent('}', '', '')
        footer.normal(`exports->Set(Nan::New("${name}").ToLocalChecked(), Nan::New<v8::FunctionTemplate>(TA_FUNC_${name})->GetFunction());`)
    })

    footer
        .normal('v8::Local<v8::Object> MATypes = Nan::New<v8::Object>();')
        .normal(...Object.keys(TA_MATypes).map((value) => {
            const [, , name] = TA_MATypes[value].split('_')
            return `MATypes->Set(Nan::New("${name}").ToLocalChecked(), Nan::New<v8::Number>(${value}));`
        }))
        .normal('exports->Set(Nan::New("MATypes").ToLocalChecked(), MATypes);')
        .undent('}')
        .normal('')
        .normal('NODE_MODULE(talib_binding, Init)', '')
    fs.writeFileSync(__dirname + '/talib-binding.generated.cc', body.build() + footer.build())
}

export function generateTypes() {
    const config: TaFuncApiXml = require('./ta_func_api.generated.json')
    const body = new ContentBuilder()
    body
        .normal(
            '/*!',
            ' * This file is generated by generate.ts, do not edit it.',
            ' */',
            '',
        )
        .normal('export declare enum MATypes {')
        .indent(...Object.keys(TA_MATypes).map((value) => {
            const [, , name] = TA_MATypes[value].split('_')
            return `${name} = ${value},`
        }))
        .undent('}')
        .normal('')
        .normal('export declare interface Record {')
        .indent('Time: number;')
        .normal('Open: number;')
        .normal('High: number;')
        .normal('Low: number;')
        .normal('Close: number;')
        .normal('Volume: number;')
        .undent('}', '')

    config.FinancialFunctions.FinancialFunction.forEach((func) => {
        const name = func.Abbreviation[0]
        const required: RequiredInputArgument[] = func.RequiredInputArguments
            ? func.RequiredInputArguments.reduce((prev, curr) => {
                return curr.RequiredInputArgument ? prev.concat(curr.RequiredInputArgument) : prev
            }, [] as RequiredInputArgument[]) : []
        const optional: OptionalInputArgument[] = func.OptionalInputArguments
            ? func.OptionalInputArguments.reduce((prev, curr) => {
                return curr.OptionalInputArgument ? prev.concat(curr.OptionalInputArgument) : prev
            }, [] as OptionalInputArgument[]) : []
        const output: OutputArgument[] = func.OutputArguments
            ? func.OutputArguments.reduce((prev, curr) => {
                return curr.OutputArgument ? prev.concat(curr.OutputArgument) : prev
            }, [] as OutputArgument[]) : []
        const doubleRequired = required.filter((argv) => RecordFields.indexOf(argv.Type[0]) === -1)
        body
            .normal(
                '/**',
                ` * ${name} - ${func.ShortDescription[0]}`,
                ' *',
            )
            .normal(...required.map((argv) => {
                return ` * @param {number[]} ${inName(argv)} - ${argv.Type[0]}`
            }))
            .normal(...optional.map((argv) => {
                return ` * @param {${optType(argv, true)}} [${optName(argv)}=${optDft(argv, true)}] - ${argv.ShortDescription[0]}`
            }))
            .normal(
                ' * @param {number} [startIdx=0] - The start index to process',
                ' * @param {number} [endIdx=inLength-1] - The end index to process, please not that the value is included, default is the input records length - 1',
            )
            .normal(
                output.length === 1
                    ? ` * @returns {number[]} - ${output[0].Name[0]} (${output[0].Type[0]})`
                    : ` * @returns {[${output.map(() => 'number[]').join(', ')}]} - [${output.map((argv) => argv.Name[0]).join(', ')}]`
            )
            .normal(' */')
            .normal(
                `export declare function ${name}(`
                + required.map((argv) => inName(argv) + ': number[]').join(', ')
                + (required.length > 0 ? ', ' : '')
                + optional.map((argv) => optName(argv) + '?: ' + optType(argv, true)).join(', ')
                + (optional.length > 0 ? ', ' : '')
                + 'startIdx?: number, endIdx?: number): '
                + (output.length === 1 ? 'number[]' : `[${output.map(() => 'number[]').join(', ')}]`)
                + ';'
            )
            .normal(
                '/**',
                ` * ${name} - ${func.ShortDescription[0]}`,
                ' *',
                ' * @param {Record[]} inRecords - The records to extract data',
            )
            .normal(...doubleRequired.map((argv) => {
                return ` * @param {string} ${inName(argv)}Name - The field name to extract from \`inRecords\``
            }))
            .normal(...optional.map((argv) => {
                return ` * @param {${optType(argv, true)}} [${optName(argv)}=${optDft(argv, true)}] - ${argv.ShortDescription[0]}`
            }))
            .normal(
                ' * @param {number} [startIdx=0] - The start index to process',
                ' * @param {number} [endIdx=inLength - 1] - The end index to process, please not that the value is included, default is the input records length - 1',
            )
            .normal(
                output.length === 1
                    ? ` * @returns {number[]} - ${output[0].Name[0]} (${output[0].Type[0]})`
                    : ` * @returns {[${output.map(() => 'number[]').join(', ')}]} - [${output.map((argv) => argv.Name[0]).join(', ')}]`
            )
            .normal(' */')
            .normal(
                `export declare function ${name}(`
                + 'inRecords: Record[], '
                + doubleRequired.map((argv) => inName(argv) + 'Name: string, ').join('')
                + optional.map((argv) => optName(argv) + '?: ' + optType(argv, true) + ', ').join('')
                + 'startIdx?: number, endIdx?: number): '
                + (output.length === 1 ? 'number[]' : `[${output.map(() => 'number[]').join(', ')}]`)
                + ';'
            )
            .normal('')
    })
    fs.writeFileSync(__dirname + '/talib-binding.generated.d.ts', body.build())
}

export function generateGyp() {
    const sources = G.sync('ta-lib/c/src/{ta_abstract,ta_common,ta_func}/**/!(ta_java_defs|excel_glue).c', {
        cwd: __dirname + '/..',
        nodir: true,
    })
    const includes = G.sync('ta-lib/c/{include,src/{ta_abstract,ta_common,ta_func}/**}/', {
        cwd: __dirname + '/..',
        nodir: false,
    })
    sources.push('src/talib-binding.generated.cc')
    includes.push('<!(node -e "require(\'nan\')")')
    const config: any = {
        targets: [
            {
                target_name: 'talib_binding',
                sources: sources,
                include_dirs: includes,
            }
        ]
    }
    fs.writeFileSync(__dirname + '/../binding.gyp', JSON.stringify(config, void 0, 2))
}

autorun(module)
