

for p in "${patterns[@]}"; do
  echo "Searching for: $p"

  grep -Rin --color=always "$p" . \
    || echo "No matches in plain files for $p"

  find . -type f -name "*.gz" -exec zgrep -in --color=always "$p" {} + \
    || echo "No matches in .gz files for $p"

  echo "--------------------------------------"
done > /Home/filterResults.txt
