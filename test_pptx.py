import datetime

import jinja2

from pptxtpl.PptxDocument import PptxDocument


data = {"product": "Pptx-tpl", "version": "1.0.0"}


jinja_env = jinja2.Environment()
jinja_env.globals["now"] = datetime.datetime.now

doc = PptxDocument("sample/sample.pptx")
doc.render(data, jinja_env)
doc.save("sample/output.pptx")
