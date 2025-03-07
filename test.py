import json
import matplotlib.font_manager

x = sorted(set(f.name for f in matplotlib.font_manager.fontManager.ttflist))

# convert to Json
json_str = json.dumps(x)
# displaying
print(type(json_str))
print("Json List:", json_str)

with open("fonts.json", "w") as final:
	json.dump(x, final)