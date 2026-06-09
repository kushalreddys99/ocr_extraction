#!/bin/bash
curl -fsSL https://ollama.com/install.sh | sh
npm install
ollama run qwen2.5vl:7b



# command to run:
# chmod +x setup.sh
# ./setup.sh