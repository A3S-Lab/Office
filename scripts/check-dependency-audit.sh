#!/usr/bin/env bash
set -euo pipefail

# Fails on any high or critical advisory except the reviewed ones below.
# Remove an entry as soon as its parent package ships a fixed range.
#
# image-size (via pptxgenjs): only pptxgenjs's Node file-path branch loads it;
# the browser bundle this package ships never references it.
reviewed_unreachable=(
  GHSA-5p2g-fcmc-qvqq
  GHSA-w3rx-r6r6-pgpr
)

ignore_arguments=()
for advisory in "${reviewed_unreachable[@]}"; do
  ignore_arguments+=("--ignore=${advisory}")
done

bun audit --audit-level=high "${ignore_arguments[@]}"
