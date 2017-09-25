#!/usr/bin/env bash
#
# build.sh
# @author acrazing
# @since 2017-09-25 23:05:26
# @desc build.sh
#

set -xe
cd ta-lib
cmake ./
make -j4
ts-node ./src/generate.ts generate-bindings
ts-node ./src/generate.ts generate-types
node-gyp configure build -j 4
