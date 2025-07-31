#!/usr/bin/env python3

import re

def extract_urls_from_text(text):
    """Extract URLs from text using regex patterns, handling line breaks."""
    
    # First, fix URLs that are split across lines by rejoining hyphenated words
    # Look for patterns like "system-\nprompt-override" and join them
    text = re.sub(r'-\s*\n\s*', '-', text)
    
    # More careful URL joining - only join if the next line looks like a URL continuation
    # Look for URL parts that end with common URL characters and continue on next line
    text = re.sub(r'(https?://[^\s]*[/\-])\s*\n\s*([a-zA-Z0-9\-/\.]+[/\.]?)', r'\1\2', text)
    
    url_patterns = [
        r'https?://[^\s<>"{}|\\^`\[\]]+',  # Standard URLs
        r'www\.[^\s<>"{}|\\^`\[\]]+',      # www URLs
        r'doi:[^\s<>"{}|\\^`\[\]]+',       # DOI URLs
    ]
    
    urls = []
    for pattern in url_patterns:
        urls.extend(re.findall(pattern, text, re.IGNORECASE))
    
    # Clean up URLs (remove trailing punctuation)
    cleaned_urls = []
    for url in urls:
        # Remove trailing punctuation
        url = re.sub(r'[.,;:!?)\]]*$', '', url)
        if url.startswith('www.'):
            url = 'https://' + url
        cleaned_urls.append(url)
    
    return cleaned_urls

# Test with the exact format from user's sample
test_text = """Zang, Guangshuo. 2025. "System Prompt Override Plugin." promptfoo. July 28. Accessed
July 28, 2025. https://www.promptfoo.dev/docs/red-team/plugins/system-
prompt-override/.

Zhang, Shuning. 2024. "Ghost of the past": Identifying and Resolving Privacy Leakage
of LLM's Memory Through Proactive User Interaction." arXiv. Oct 19. Accessed
July 30, 2025. https://arxiv.org/html/2410.14931v1."""

print("Testing URL extraction:")
print("Original text with line breaks:")
print(repr(test_text))
print()

urls = extract_urls_from_text(test_text)
print(f'Found {len(urls)} URLs:')
for i, url in enumerate(urls):
    print(f'  {i+1}: {url}')