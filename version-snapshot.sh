#!/usr/bin/env bash

version=$(sed -n -E 's/^version\s*=\s*(.+)$/\1/p' gradle.properties)
if [[ -z $version ]]; then
  echo '[ERROR] version not found' >&2
  exit 1
fi
if [[ ! $version =~ ^([[:digit:]]+)\.([[:digit:]]+)\.([[:digit:]]+)(-SNAPSHOT)?$ ]]; then
  echo '{ERROR] version must be 0.0.0 or 0.0.0-SNAPSHOT' >&2
  exit 1
fi
[[ -n ${BASH_REMATCH[4]} ]] && exit 0
major_v=${BASH_REMATCH[1]}
minor_v=${BASH_REMATCH[2]}
patch_v=$((BASH_REMATCH[3] + 1))
new_version="$major_v.$minor_v.$patch_v-SNAPSHOT"
sed -i -E "s/^version\s*=.+$/version = $new_version/" gradle.properties
