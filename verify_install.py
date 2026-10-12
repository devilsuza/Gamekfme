from pathlib import Path
import ast, py_compile, sys
p=Path(__file__).with_name("app.py")
py_compile.compile(str(p),doraise=True)
tree=ast.parse(p.read_text(encoding="utf-8"))
routes=[]
for n in ast.walk(tree):
    if isinstance(n,(ast.FunctionDef,ast.AsyncFunctionDef)):
        for d in n.decorator_list:
            if isinstance(d,ast.Call) and isinstance(d.func,ast.Attribute) and d.func.attr=="route" and d.args:
                try: routes.append(ast.literal_eval(d.args[0]))
                except Exception: pass
required=["/","/login","/apply_leave","/export","/developer/portal_update","/developer/recovery","/api/build"]
missing=[x for x in required if x not in routes]
duplicates=sorted({x for x in routes if routes.count(x)>1})
print("Python syntax: PASS")
print("Routes:",len(routes))
print("Missing:",missing)
print("Duplicates:",duplicates)
if missing or duplicates: sys.exit(1)
print("Package verification: PASS")
