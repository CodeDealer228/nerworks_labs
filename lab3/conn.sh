#!/bin/bash
P=${1:-22}
ss -Htn state established |
awk -v p=":$P" '$3 ~ p"$" {print $4}' |
sed 's/:[0-9]*$//' | sort | uniq -c |
sort -rn
