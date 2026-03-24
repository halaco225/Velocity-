with open('public/index.html', 'r', encoding='utf-8') as f:
    content = f.read()

marker_before = '>IST<19%</th>'
marker_after = '  </tr>`'

idx_before = content.find(marker_before)
idx_after = content.find(marker_after, idx_before)

print(f"Before marker at: {idx_before}")
print(f"After marker at: {idx_after}")

bad_section = content[idx_before + len(marker_before):idx_after]
print("Current bad section:")
print(repr(bad_section))