import os

NOTES_DIR = "notes"
OUTPUT_FILE = "INDEX.md"


def get_title(filepath):
	with open(filepath, "r", encoding="utf-8") as f:
		first_line = f.readline().strip()
	if first_line.startswith('#'):
		return first_line.lstrip('#').strip()
	else:
		return filepath

def main():
	entries = []
	for filename in sorted(os.listdir(NOTES_DIR)):
		if filename.endswith(".md"):
			filepath = os.path.join(NOTES_DIR, filename)
			title = get_title(filepath)
			entries.append(f"- [{title}]({NOTES_DIR}/{filename})")
	with open(OUTPUT_FILE, "w", encoding="utf-8") as f:
		f.write("# Index of Notes\n\n")
		f.write("\n".join(entries))
		f.write("\n")

	print(f"Generated {OUTPUT_FILE} with {len(entries)} notes.")
if __name__ == "__main__":
	main()
