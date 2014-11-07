#!/bin/bash 

rm -rf test/spreadsheets/contents
mkdir test/spreadsheets/contents

for f in test/spreadsheets/*.xlsx; do
  filterEnd=${f%.xlsx}
  filterBeg=${filterEnd#"test/spreadsheets/"}

  if [[ ! $filterBeg =~ ^('~''$') ]]; then
  unzip $f -d test/spreadsheets/contents/"$filterBeg"_contents
  fi

done
