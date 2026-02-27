import json, re

data = json.load(open('cia_rdp96.json'))
keyword = "salmonella"                          # kellman; shafran; galland
matches = []

for date, titles in data.items():
    for title, content_dict in titles.items():
        content = str(content_dict)  # Convert dict to string
        if re.search(keyword, content, re.I):
            matches.append({"date": date, "title": title, "content": content})

json.dump(matches, open('results.json', 'w'), indent=2)
print(f"Saved {len(matches)} matches")