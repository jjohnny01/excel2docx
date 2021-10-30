#!/bin/bash

uid=$(id -u)

if [ "$uid" -ne 0 ]
    then echo "Error: To install, please run as root (uid 0)."
    exit
fi

cp ./excel2docx.py ./excel2docx
mv ./excel2docx /usr/local/bin/

echo "Installation complete. Usage: $ excel2docx"
