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

interface Json2ts {
    convertObjectToTsInterfaces<T>(content: T, name?: string): string;
    toUpperFirstLetter(text: string): string;
    toLowerFirstLetter(text: string): string;
    removeMajority(text: string): string;
}

const json2ts: Json2ts = require('json2ts')
json2ts.toLowerFirstLetter = json2ts.toUpperFirstLetter = json2ts.removeMajority = (text) => text

export function generateTypes() {
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

autorun(module)
