#!/bin/bash

for test in *test_*.py
do 
  echo "Running $test"
  python2 $test
done
