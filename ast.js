const files = require('fs');
const types = require("@babel/types");
const parser = require("@babel/parser");
const template = require("@babel/template").default;
const traverse = require("@babel/traverse").default;
const generator = require("@babel/generator").default;
const NodePath = require("@babel/traverse").NodePath;


class XiaoHongShu {
    constructor(file_path) {
        this.ast = parser.parse(files.readFileSync(file_path, "utf-8"));
        this.stringPool = [];
        this.stringPoolList = null;
        this.countValue = null;
        this.otr = null;
        this.transFunc = null;
        this.transFuncName = null;
        this.transCode = null;
    }

    save_file() {
        const {code: newCode} = generator(this.ast);
        files.writeFileSync(
            './xs/decode.js',
            newCode,
            "utf-8"
        );
    }

    fix_code() {
        let objn = {};
        traverse(this.ast, {
            Literal(path) {
                let {parentPath, node} = path;
                if (!types.isNumericLiteral(node)) return;
                if (!types.isObjectProperty(parentPath.node)) return;
                if (types.isCallExpression(parentPath.parentPath.parent)) return;
                if (parentPath.parentPath.parent.id === undefined) return;
                let objNumLit = parentPath.parentPath.parent.id.name;
                let objs = {};
                parentPath.parentPath.parent.init.properties.forEach(res => {
                    if (!types.isNumericLiteral(res.value)) return;
                    objs[res.key.name] = res.value.extra.raw
                });
                objn[objNumLit] = objs;
            }
        });
        traverse(this.ast, {
            MemberExpression: (path) => {
                let {object, property} = path.node;
                if (!types.isIdentifier(object)) return;
                if (!types.isIdentifier(property)) return;
                let objName = object.name;
                if (!objn[objName]) return;
                let objValue = objn[objName][property.name];
                let num = parseInt(objValue, 16);
                if (objValue === undefined) return;
                const node = types.numericLiteral(num);
                node.extra = { raw: objValue, rawValue: num };
                path.replaceWith(node);
            }
        });
        Object.keys(objn).forEach((key) => {
            traverse(this.ast, {
                VariableDeclarator: (path) => {
                    let {id, init} = path.node;
                    if (!types.isIdentifier(id)) return;
                    if (!types.isObjectExpression(init)) return;
                    if (id.name !== key) return;
                    path.remove()
                }
            })
        });
    }

    convert_code() {
        let fn, sfn, func, strFunc, initFunc, initFuncName = 'g';
        traverse(this.ast, {
            CallExpression: (path) => {
                let {callee, arguments: args} = path.node;
                if (!types.isFunctionExpression(callee)) return;
                if (args.length !== 2) return;
                let {id, params, body} = callee;
                fn = body.body[0].declarations[0].init.name;
                this.transFuncName = fn;
                path.stop()
            }
        });
        traverse(this.ast, {
            FunctionDeclaration: (path) => {
                let {id, params, body} = path.node;
                if (id.name !== fn) return;
                if (params.length === 0) return;
                func = generator(path.node).code;
                this.transFunc = func;
                sfn = body.body[0].declarations[0].init.callee.name;
                let dfr = body.body[1].expression.right.body.body[0].expression.right;
                this.countValue = dfr.right.value;
                this.otr = dfr.operator
            }
        });
        traverse(this.ast, {
            FunctionDeclaration: (path) => {
                let {id, params, body} = path.node;
                if (id.name !== sfn) return;
                strFunc = generator(path.node).code;
                const ay = body.body[0].declarations[0].init.elements;
                ay.forEach(res => {
                    this.stringPool.push(res.value)
                });
                path.remove()
            }
        });
        // init
        traverse(this.ast, {
            CallExpression: (path) => {
                let {callee, arguments: args} = path.node;
                if (!types.isFunctionExpression(callee)) return;
                if (args.length !== 2) return;
                let {id, params, body} = callee;
                if (id !== null) return;
                if (params.length === 0) return;
                if (!types.isBlockStatement(body)) return;
                if (!types.isWhileStatement(body.body[body.body.length - 1])) return;
                path.node.callee.id = types.identifier(initFuncName);
                const retValue = body.body[1].declarations[body.body[0].declarations.length - 1].id.name;
                const returnNode = types.returnStatement(types.identifier(retValue));
                const run = generator(types.expressionStatement(types.callExpression(types.identifier(initFuncName), args))).code;
                body.body.push(returnNode);
                const f = types.functionDeclaration(path.node.callee.id, params, body);
                initFunc = generator(f).code;
                initFunc = func + '\n' + strFunc + '\n' + initFunc + '\n' + run;
                this.stringPool = eval(initFunc);
                this.stringPoolList = JSON.stringify(this.stringPool);
                this.transCode = `function ${sfn}(){return ${this.stringPoolList}}` + '\n' + this.transFunc + '\n' + this.transFuncName;
                path.remove()
            }
        })
    }

    trans_code() {
        traverse(this.ast, {
            CallExpression: (path) => {
                let {callee, arguments: args} = path.node;
                if (!types.isIdentifier(callee)) return;
                if (!types.isNumericLiteral(args[0])) return;
                if (args[0].extra === undefined) return;
                const codeValue = args[0].extra.raw;
                if (!(codeValue.includes('0x'))) return;
                const code = eval(this.transCode + `(${args[0].value})`);
                if (code === undefined) return;
                const decode = types.stringLiteral(code);
                path.replaceWith(decode);
            }
        })
    }


    start() {
        this.fix_code();
        this.convert_code();
        this.trans_code();
        this.save_file();
    }

}

console.time('处理完毕，耗时');

let xhs_ast = new XiaoHongShu('./xs.js');
xhs_ast.start();


console.timeEnd('处理完毕，耗时');