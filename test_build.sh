#!/bin/bash

echo "🔧 Testing Resume Builder..."
echo ""

# Test different role variants
roles=("python" "java" "fullstack" "backend" "cloud" "all")

for role in "${roles[@]}"
do
    echo "📝 Generating resume for: $role"
    python build_docx.py "$role"
    echo ""
done

echo "✅ All resumes generated successfully!"
echo "📂 Check the outputs/ directory for results"