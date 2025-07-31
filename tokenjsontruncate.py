import json

with open("token.json") as f:
    token_data = f.read()
    escaped = json.dumps(token_data)  # This escapes quotes, newlines, etc
    print(escaped)
