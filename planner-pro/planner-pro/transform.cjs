const fs = require("fs");
const parser = require("@babel/parser");
const generate = require("@babel/generator").default;
const traverse = require("@babel/traverse").default;
const t = require("@babel/types");

const src = fs.readFileSync(process.argv[2], "utf8");

let ast;
try {
  ast = parser.parse(src, {
    plugins: ["jsx"],
    sourceType: "module"
  });
} catch(e) {
  console.error("PARSE ERROR:", e.message);
  process.exit(1);
}

// ── 1. Supprimer les imports ES module ─────────────────────────────────────
traverse(ast, {
  ImportDeclaration(path) {
    path.remove();
  },
  ExportDefaultDeclaration(path) {
    // export default function App() {} → function App() {}
    const decl = path.node.declaration;
    path.replaceWith(decl);
  },
  ExportNamedDeclaration(path) {
    if (path.node.declaration) {
      path.replaceWith(path.node.declaration);
    } else {
      path.remove();
    }
  }
});

// ── 2. Transformer JSX ────────────────────────────────────────────────────
function jsxNameToNode(name) {
  if (t.isJSXMemberExpression(name)) {
    return t.memberExpression(
      jsxNameToNode(name.object),
      t.identifier(name.property.name)
    );
  }
  if (t.isJSXNamespacedName(name)) {
    return t.stringLiteral(name.namespace.name + ":" + name.name.name);
  }
  const n = name.name;
  if (n[0] === n[0].toUpperCase() && n[0] !== n[0].toLowerCase()) {
    return t.identifier(n);
  }
  return t.stringLiteral(n);
}

function jsxAttrValue(val) {
  if (!val) return t.booleanLiteral(true);
  if (t.isJSXExpressionContainer(val)) {
    return t.isJSXEmptyExpression(val.expression) ? t.booleanLiteral(true) : val.expression;
  }
  return val; // StringLiteral
}

function jsxChild(child) {
  if (t.isJSXText(child)) {
    const val = child.value.replace(/\r\n|\n|\r/g, "\n");
    // Collapse whitespace like React does
    const lines = val.split("\n");
    const parts = lines.map((line, i) => {
      if (i === 0 && i === lines.length - 1) return line;
      if (i === 0) return line.trimEnd();
      if (i === lines.length - 1) return line.trimStart();
      return line.trim();
    }).filter(l => l !== "");
    if (parts.length === 0) return null;
    return t.stringLiteral(parts.join(" "));
  }
  if (t.isJSXExpressionContainer(child)) {
    if (t.isJSXEmptyExpression(child.expression)) return null;
    return child.expression;
  }
  if (t.isJSXSpreadChild(child)) {
    return t.spreadElement(child.expression);
  }
  return child; // JSXElement / JSXFragment - transformed by visitor
}

traverse(ast, {
  JSXFragment(path) {
    const children = path.node.children.map(jsxChild).filter(Boolean);
    path.replaceWith(t.callExpression(
      t.memberExpression(t.identifier("React"), t.identifier("createElement")),
      [t.memberExpression(t.identifier("React"), t.identifier("Fragment")), t.nullLiteral(), ...children]
    ));
  },
  JSXElement(path) {
    const opening = path.node.openingElement;
    const tagNode = jsxNameToNode(opening.name);

    let propsNode;
    if (opening.attributes.length === 0) {
      propsNode = t.nullLiteral();
    } else {
      const hasSpread = opening.attributes.some(a => t.isJSXSpreadAttribute(a));
      if (hasSpread) {
        // Object.assign({}, ...attrs)
        const parts = [];
        let curr = [];
        opening.attributes.forEach(attr => {
          if (t.isJSXSpreadAttribute(attr)) {
            if (curr.length) { parts.push(t.objectExpression(curr)); curr = []; }
            parts.push(attr.argument);
          } else {
            curr.push(t.objectProperty(
              t.identifier(t.isJSXNamespacedName(attr.name) ? attr.name.namespace.name + "_" + attr.name.name.name : attr.name.name),
              jsxAttrValue(attr.value)
            ));
          }
        });
        if (curr.length) parts.push(t.objectExpression(curr));
        propsNode = t.callExpression(
          t.memberExpression(t.identifier("Object"), t.identifier("assign")),
          [t.objectExpression([]), ...parts]
        );
      } else {
        const props = opening.attributes.map(attr => {
          const key = t.isJSXNamespacedName(attr.name)
            ? attr.name.namespace.name + ":" + attr.name.name.name
            : attr.name.name;
          return t.objectProperty(
            /^[a-zA-Z_$][a-zA-Z0-9_$]*$/.test(key) ? t.identifier(key) : t.stringLiteral(key),
            jsxAttrValue(attr.value)
          );
        });
        propsNode = t.objectExpression(props);
      }
    }

    const children = path.node.children.map(jsxChild).filter(Boolean);
    path.replaceWith(t.callExpression(
      t.memberExpression(t.identifier("React"), t.identifier("createElement")),
      [tagNode, propsNode, ...children]
    ));
  }
});

// ── 3. Générer le code JS final ──────────────────────────────────────────
const output = generate(ast, { comments: false, compact: false }).code;
fs.writeFileSync(process.argv[3], output);
console.log("OK - " + output.length + " chars");
