import pandas as pd
import requests
import re
import random
import json
import time
import os
from datetime import datetime
from colorama import init, Fore, Style
from anthropic import Anthropic  # Import the official Anthropic Python client

# Initialize colorama for colored terminal output
try:
    init()
    COLORAMA_AVAILABLE = True
except ImportError:
    class DummyColor:
        def __getattr__(self, name):
            return ""
    COLORAMA_AVAILABLE = False
    Fore = DummyColor()
    Style = DummyColor()
    print("Colorama not installed. Output will not be colored.\nTry running: pip install colorama")

### === CONFIG SECTION === ###
excel_file = 'products.xlsx'

# WooCommerce API configuration (credentials preserved as provided)
WOOCOMMERCE_API_URL = "https://xsellpoint.com/wp-json/wc/v3/products"
WOOCOMMERCE_USER = "ck_35a17db87828f5bd7733ad0968562e2dd1d160bf"
WOOCOMMERCE_PASS = "cs_4afc575bd0e557c2ee374b5c4253f0bfd6e80980"

# Claude API configuration - Using Anthropic's Python client
CLAUDE_API_KEY = "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"  # Replace with your actual Claude API key

# WordPress media configuration
WP_MEDIA_URL = "https://xsellpoint.com/wp-json/wp/v2/media"
wp_media_username = 'alaa'
wp_media_password = 'QirZ 451o Y9iC 3FIl L2YZ nLDb'
### ======================== ###

# Track failures for summary
failed_products = []

def print_colored(text, color=Fore.WHITE, is_bold=False, end='\n'):
    bold_style = Style.BRIGHT if is_bold else ""
    print(f"{bold_style}{color}{text}{Style.RESET_ALL}", end=end)

def read_products_from_excel(file_path):
    """Read product data from an Excel file."""
    try:
        if not os.path.exists(file_path):
            print_colored(f"Error: Excel file '{file_path}' not found.", Fore.RED, True)
            return pd.DataFrame()
        df = pd.read_excel(file_path)
        print_colored(f"Read {len(df)} products from Excel file.", Fore.GREEN)
        print_colored("Excel file contains these columns:", Fore.CYAN)
        for col in df.columns:
            print_colored(f"  - {col}", Fore.CYAN)
        df.columns = [col.lower() for col in df.columns]
        id_columns = [col for col in df.columns if col in ['id', 'product id', 'product_id', 'productid']]
        if id_columns:
            df = df.rename(columns={id_columns[0]: 'id'})
            print_colored(f"Using '{id_columns[0]}' as the product ID column", Fore.GREEN)
        else:
            print_colored("Warning: No ID column found in Excel.", Fore.YELLOW, True)
            for col in df.columns:
                print_colored(f"  - {col}", Fore.YELLOW)
        return df
    except Exception as e:
        print_colored(f"Error reading Excel file: {e}", Fore.RED, True)
        return pd.DataFrame()

def clean_html_content(content):
    """Clean HTML content by removing code block markers and ensuring proper HTML structure."""
    if not content:
        return content
    
    # Remove markdown code block markers and other markdown formatting
    content = re.sub(r'```html\s*', '', content)
    content = re.sub(r'```\s*', '', content)
    content = re.sub(r'\*\*\*(.*?)\*\*\*', '<strong>\\1</strong>', content)  # Replace *** with <strong>
    content = re.sub(r'\*\*(.*?)\*\*', '<strong>\\1</strong>', content)  # Replace ** with <strong>
    content = re.sub(r'\*(.*?)\*', '<em>\\1</em>', content)  # Replace * with <em>
    
    # Ensure proper HTML structure
    if not re.search(r'^\s*<\w+', content):
        content = f"<p>{content}</p>"
    
    # Replace double line breaks with paragraph tags
    content = re.sub(r'\n\s*\n', '</p><p>', content)
    
    # Make sure there are no leftover markdown formatting artifacts
    content = re.sub(r'`([^`]+)`', '<code>\\1</code>', content)
    
    return content

def fetch_additional_info(url):
    """Fetch additional product information from the provided URL."""
    if not url or pd.isna(url):
        return "No URL provided"
    
    try:
        print_colored(f"Fetching additional info from: {url}", Fore.BLUE)
        response = requests.get(url, timeout=10)
        if response.status_code == 200:
            # Just return the first 1000 characters for now
            # Could be enhanced with proper HTML parsing if needed
            text = response.text[:5000]
            print_colored(f"Successfully fetched {len(text)} characters of additional info", Fore.GREEN)
            return text
        else:
            print_colored(f"Failed to fetch info: Status code {response.status_code}", Fore.RED)
            return f"Failed to fetch info: Status code {response.status_code}"
    except Exception as e:
        print_colored(f"Error fetching additional info: {e}", Fore.RED)
        return f"Error fetching additional info: {e}"

def create_optimized_permalink(product_name, focus_keyword, max_length=65):
    """
    Create an SEO-friendly permalink that starts with the focus keyword.
    Ensures the permalink is no longer than max_length characters.
    """
    # Replace spaces with hyphens and remove any special characters
    simplified_keyword = focus_keyword.lower()
    simplified_keyword = re.sub(r'[^a-z0-9\s]', '', simplified_keyword)
    simplified_keyword = re.sub(r'\s+', '-', simplified_keyword.strip())
    
    # Create permalink starting with focus keyword
    base_permalink = simplified_keyword
    
    # Add product name if there's space (after removing special characters)
    if len(base_permalink) < max_length - 5:
        simplified_name = re.sub(r'[^a-z0-9\s]', '', product_name.lower())
        simplified_name = re.sub(r'\s+', '-', simplified_name.strip())
        
        if simplified_name and simplified_name != simplified_keyword:
            base_permalink += f"-{simplified_name}"
    
    # Ensure it's not too long
    if len(base_permalink) > max_length:
        base_permalink = base_permalink[:max_length]
        
    # Remove trailing hyphens
    base_permalink = base_permalink.rstrip('-')
    
    return base_permalink

def validate_seo_content(product_name, sections):
    """Validate that all required SEO content sections are present."""
    required_fields = ['LONG DESCRIPTION', 'SHORT DESCRIPTION', 'META TITLE', 'META DESCRIPTION', 'FOCUS KEYWORD', 'SECONDARY KEYWORDS', 'TAGS', 'PERMALINK']
    missing = [field for field in required_fields if not sections.get(field)]
    if missing:
        print_colored(f"❌ Missing required fields: {', '.join(missing)}", Fore.RED, True)
        error_detail = {
            "product_name": product_name,
            "error": f"Missing required fields: {', '.join(missing)}",
            "details": "API response did not contain all required sections"
        }
        failed_products.append(error_detail)
        return False
    return True

def calculate_seo_score(seo_content):
    """
    Calculate an SEO score based on RankMath criteria.
    Score components:
      - Meta title starts with focus keyword (25 pts)
      - Contains a power word (15 pts)
      - Contains a number (10 pts)
      - Meta title <= 60 chars (10 pts)
      - Permalink starts with focus keyword (15 pts)
      - Permalink <= 65 chars (5 pts)
      - Meta description length between 140-155 (10 pts)
      - Meta description contains focus keyword (10 pts)
    """
    score = 0
    focus = seo_content.get("focus_keyword", "").lower()
    meta_title = seo_content.get("meta_title", "")
    meta_desc = seo_content.get("meta_description", "")
    permalink = seo_content.get("permalink", "").lower()
    power_words = ["exclusive", "premium", "luxurious", "ultimate", "authentic", "stunning"]

    # Check if meta title starts with focus keyword (critical for RankMath)
    if meta_title.lower().startswith(focus):
        score += 25
        print_colored("✓ Meta title starts with focus keyword", Fore.GREEN)
    else:
        print_colored("✗ Meta title doesn't start with focus keyword", Fore.RED)

    # Check for power words
    if any(pw in meta_title.lower() for pw in power_words):
        score += 15
        print_colored("✓ Meta title contains a power word", Fore.GREEN)
    else:
        print_colored("✗ Meta title doesn't contain a power word", Fore.RED)

    # Check for numbers
    if re.search(r'\d', meta_title):
        score += 10
        print_colored("✓ Meta title contains a number", Fore.GREEN)
    else:
        print_colored("✗ Meta title doesn't contain a number", Fore.RED)

    # Check meta title length
    if len(meta_title) <= 60:
        score += 10
        print_colored(f"✓ Meta title length is good: {len(meta_title)}/60", Fore.GREEN)
    else:
        print_colored(f"✗ Meta title too long: {len(meta_title)}/60", Fore.RED)

    # Check if permalink starts with focus keyword
    focus_in_permalink = focus.replace(' ', '-')
    if permalink.startswith(focus_in_permalink):
        score += 15
        print_colored("✓ Permalink starts with focus keyword", Fore.GREEN)
    else:
        print_colored("✗ Permalink doesn't start with focus keyword", Fore.RED)

    # Check permalink length
    if len(permalink) <= 65:
        score += 5
        print_colored(f"✓ Permalink length is good: {len(permalink)}/65", Fore.GREEN)
    else:
        print_colored(f"✗ Permalink too long: {len(permalink)}/65", Fore.RED)

    # Check meta description length
    if 140 <= len(meta_desc) <= 155:
        score += 10
        print_colored(f"✓ Meta description length is ideal: {len(meta_desc)}", Fore.GREEN)
    else:
        print_colored(f"✗ Meta description length is not ideal: {len(meta_desc)}", Fore.RED)

    # Check if meta description contains focus keyword
    if focus in meta_desc.lower():
        score += 10
        print_colored("✓ Meta description contains focus keyword", Fore.GREEN)
    else:
        print_colored("✗ Meta description doesn't contain focus keyword", Fore.RED)

    return score

def create_seo_title(focus_keyword, product_name, max_length=60):
    """
    Create an SEO-optimized title that always begins with the focus keyword.
    Includes a power word and a number to boost SEO score.
    """
    # Ensure focus keyword is capitalized properly
    focus_keyword = focus_keyword.strip()
    focus_capitalized = ' '.join(word.capitalize() for word in focus_keyword.split())
    
    # Select a power word and a sentiment word
    power_words = ["Exclusive", "Premium", "Luxury", "Authentic", "Ultimate", "Elegant"]
    sentiment_words = ["Experience", "Collection", "Selection", "Choice", "Quality", "Perfection"]
    
    power_word = random.choice(power_words)
    sentiment_word = random.choice(sentiment_words)
    number = random.randint(1, 99)
    
    # Create the title format with focus keyword at the beginning
    seo_title = f"{focus_capitalized} - {power_word} {number}% {sentiment_word}"
    
    # If it's too long, simplify
    if len(seo_title) > max_length:
        seo_title = f"{focus_capitalized} - {power_word} {number}"
    
    # Still too long? Use just the focus keyword
    if len(seo_title) > max_length:
        seo_title = focus_capitalized
        
    # Ensure title is within limits
    if len(seo_title) > max_length:
        seo_title = seo_title[:max_length]
    
    return seo_title

def improve_seo_fields(seo_content, product_name):
    """
    Improve SEO fields to ensure they meet all RankMath criteria,
    particularly ensuring the focus keyword appears at the beginning of the title.
    """
    focus_keyword = seo_content.get("focus_keyword", "").strip()
    if not focus_keyword:
        focus_keyword = product_name.split()[0]
        seo_content["focus_keyword"] = focus_keyword
    
    # Force the meta title to start with the focus keyword
    meta_title = seo_content.get("meta_title", "")
    if not meta_title.lower().startswith(focus_keyword.lower()):
        print_colored("⚠️ Rebuilding meta title to start with focus keyword", Fore.YELLOW)
        meta_title = create_seo_title(focus_keyword, product_name)
    
    # Force permalink to start with focus keyword
    permalink = seo_content.get("permalink", "")
    focus_in_permalink = focus_keyword.lower().replace(' ', '-')
    if not permalink.startswith(focus_in_permalink):
        print_colored("⚠️ Rebuilding permalink to start with focus keyword", Fore.YELLOW)
        permalink = create_optimized_permalink(product_name, focus_keyword)
    
    # Ensure meta description contains focus keyword and is the right length
    meta_desc = seo_content.get("meta_description", "")
    if focus_keyword.lower() not in meta_desc.lower():
        print_colored("⚠️ Rebuilding meta description to include focus keyword", Fore.YELLOW)
        meta_desc = f"{focus_keyword} offers an exclusive, luxurious experience. Shop now for the ultimate choice that will delight you every day!"
    
    # Adjust meta description length to be between 140-155 characters
    if len(meta_desc) < 140:
        extra_text = " Perfect for all occasions. Try it today and experience the difference!"
        meta_desc += extra_text[:140 - len(meta_desc)]
    elif len(meta_desc) > 155:
        meta_desc = meta_desc[:152] + "..."
    
    # Update the SEO content with improved fields
    seo_content["meta_title"] = meta_title
    seo_content["permalink"] = permalink
    seo_content["meta_description"] = meta_desc
    
    return seo_content

def generate_seo_content(product_name, product_description, additional_info="", brand_name=""):
    """
    Generate SEO content using the Claude API based on provided product data.
    The system prompt instructs detailed product description and SEO fields.
    Uses the official Anthropic Python client.
    """
    system_prompt = (
        "You are an expert eCommerce SEO product description writer specializing in optimizing product content. "
        "Your task is to write detailed and SEO-optimized product descriptions based on the provided information.\n\n"
        "Focus on creating content that ranks well in RankMath plugin. Critical requirements:\n"
        "- SEO Title MUST start with the Focus Keyword exactly\n"
        "- Permalink MUST start with the Focus Keyword\n"
        "- Content should be clean HTML without Markdown formatting\n\n"
        "Content Requirements:\n"
        "1. Long Description (300+ words, HTML format):\n"
        "   - Include detailed and informative content optimized for SEO\n"
        "   - Use <strong> tags for highlighting important keywords (not Markdown)\n"
        "   - Start with the Focus Keyword and repeat it naturally\n"
        "   - Include the Focus Keyword in subheadings (<h2>, <h3>, <h4>)\n"
        "   - Include a Product Information Table (Size, Color, Material, Brand Name)\n"
        "   - Include Key Features, Benefits, and overview\n"
        "   - Answer one frequently searched question related to the product\n"
        "   - Use emoticons/icons to evoke emotional connection\n"
        "   - Include 3-4 internal links to related products\n\n"
        "2. Short Description (50 words max):\n"
        "   - Concise and engaging, highlighting uniqueness and key features\n"
        "   - Provided as plain text without any Markdown formatting\n\n"
        "3. SEO Elements (Optimized for Rank Math SEO Plugin):\n"
        "   - SEO Meta Title: MUST start with the exact Focus Keyword, be under 60 characters, and include a power word and a number\n"
        "   - SEO Permalink: MUST start with the Focus Keyword and be URL-friendly\n"
        "   - Meta Description: 140-155 characters, must include the Focus Keyword, with a call to action\n"
        "   - Focus Keyword: Generate EXACTLY THREE primary focus keyword that best represents the product\n"
        "   - Secondary Keywords: Generate EXACTLY TWO secondary keywords that complement the focus keyword\n"
        "   - Tags: Generate EXACTLY THREE product tags that are relevant to the product\n\n"
        "Output MUST include these EXACT section headers in your response:\n"
        "LONG DESCRIPTION:\n"
        "SHORT DESCRIPTION:\n"
        "META TITLE:\n"
        "META DESCRIPTION:\n"
        "FOCUS KEYWORD:\n"
        "SECONDARY KEYWORDS:\n"
        "TAGS:\n"
        "PERMALINK:\n"
        "Do not include any Markdown formatting like ``` or ** in your output."
    )
    
    user_message = f"""Product Name: {product_name}
Product Description: {product_description if product_description else 'Not available'}
Brand: {brand_name if brand_name else 'Not specified'}
Additional Information: {additional_info if additional_info else 'None available'}

Generate the comprehensive SEO content following the EXACT format specified. Remember that the META TITLE and PERMALINK must start with the Focus Keyword exactly as provided. Generate exactly ONE focus keyword, TWO secondary keywords, and THREE product tags.
"""
    try:
        # Initialize the Anthropic client with your API key
        client = Anthropic(api_key=CLAUDE_API_KEY)
        
        print_colored("🚀 Sending request to Claude API using official Python SDK...", Fore.BLUE)
        start_time = datetime.now()
        
        # Make the API request using the client
        response = client.messages.create(
            model="claude-3-7-sonnet-20250219",
            max_tokens=4000,
            temperature=0.2,
            system=system_prompt,
            messages=[
                {"role": "user", "content": user_message}
            ]
        )
        
        duration = (datetime.now() - start_time).total_seconds()
        print_colored(f"⏱️ API request took {duration:.2f} seconds", Fore.BLUE)
        
        # Extract the generated text from the response
        generated_text = response.content[0].text
        print_colored("✅ Received content from Claude API", Fore.GREEN)
        
        # Save raw response for debugging
        os.makedirs("debug", exist_ok=True)
        safe_name = re.sub(r'[^\w\-_\. ]', '_', product_name)
        with open(f"debug/{safe_name}_raw_response.txt", "w", encoding="utf-8") as f:
            f.write(generated_text)
        
        # Preview generated content (first 10 lines)
        preview = "\n".join(generated_text.split('\n')[:10])
        print_colored("📄 Content preview (first few lines):", Fore.CYAN)
        print(preview)
        
        # Parse sections from the generated text
        sections = {
            'LONG DESCRIPTION': '',
            'SHORT DESCRIPTION': '',
            'META TITLE': '',
            'META DESCRIPTION': '',
            'FOCUS KEYWORD': '',
            'SECONDARY KEYWORDS': '',
            'TAGS': '',
            'PERMALINK': ''
        }
        current_section = None
        content_buffer = []
        
        # Process line by line to extract sections
        for line in generated_text.split('\n'):
            line = line.strip()
            if not line:
                continue
                
            # Check if this line is a section header
            is_section_header = False
            for section in sections:
                header = section + ":"
                if line.upper().startswith(header) or line.upper() == section:
                    # Save current section's content before switching
                    if current_section and content_buffer:
                        sections[current_section] = "\n".join(content_buffer).strip()
                    current_section = section
                    content_buffer = []
                    is_section_header = True
                    break
            
            # If not a section header, add to current section's content
            if not is_section_header and current_section:
                content_buffer.append(line)
        
        # Save the last section's content
        if current_section and content_buffer:
            sections[current_section] = "\n".join(content_buffer).strip()
        
        # Validate that all required sections are present
        if not validate_seo_content(product_name, sections):
            return None
        
        # Clean and post-process fields
        focus_keyword = sections['FOCUS KEYWORD'].strip()
        secondary_keywords = sections['SECONDARY KEYWORDS'].strip()
        long_desc = clean_html_content(sections['LONG DESCRIPTION'])
        short_desc = clean_html_content(sections['SHORT DESCRIPTION'])
        meta_title = sections['META TITLE'].strip()
        meta_description = sections['META DESCRIPTION'].strip()
        permalink = sections['PERMALINK'].strip()
        tags = sections['TAGS'].strip()
        
        # Format keywords for RankMath
        all_keywords = f"{focus_keyword}, {secondary_keywords}"
        
        # Construct the SEO content object
        seo_content = {
            "long_description": long_desc,
            "short_description": short_desc,
            "meta_title": meta_title,
            "meta_description": meta_description,
            "focus_keyword": focus_keyword,
            "secondary_keywords": secondary_keywords,
            "all_keywords": all_keywords,
            "tags": tags,
            "permalink": permalink,
            "product_name": product_name
        }
        
        # Improve SEO fields to ensure they meet all RankMath criteria
        seo_content = improve_seo_fields(seo_content, product_name)
        
        # Calculate SEO score
        print_colored("SEO Score Breakdown:", Fore.MAGENTA, True)
        final_score = calculate_seo_score(seo_content)
        print_colored(f"Final SEO score: {final_score}/100", Fore.MAGENTA, True)
        
        # Append an image tag to the long description if not present
        if '<img' not in seo_content["long_description"].lower():
            image_tag = f'<p><img src="https://via.placeholder.com/300" alt="{focus_keyword}" style="max-width:100%;"></p>'
            seo_content["long_description"] += "\n" + image_tag
        
        return seo_content
        
    except Exception as e:
        print_colored(f"❌ Error during SEO content generation: {e}", Fore.RED, True)
        import traceback
        print_colored(traceback.format_exc(), Fore.RED)
        failed_products.append({
            "product_name": product_name,
            "error": "Exception during SEO generation",
            "details": str(e)
        })
        return None

def update_product_info(product_id, product_title, seo_content):
    """
    Update product information using the WooCommerce REST API.
    Uses the SEO content generated by the Claude API to update the product.
    Keeps the original product title and only updates meta data in RankMath.
    """
    try:
        focus_keyword = seo_content.get("focus_keyword", "")
        all_keywords = seo_content.get("all_keywords", "")
        
        # Prepare tag data from the generated tags
        tags = []
        if seo_content.get("tags"):
            tag_list = [tag.strip() for tag in seo_content.get("tags").split(",") if tag.strip()]
            if tag_list:
                tags = [{"name": tag} for tag in tag_list]
        
        # Updated data structure to keep original title
        data = {
            # Keep original product name/title
            "description": seo_content.get("long_description"),
            "short_description": seo_content.get("short_description"),
            "slug": seo_content.get("permalink"),
            "meta_data": [
                {"key": "rank_math_title", "value": seo_content.get("meta_title")},
                {"key": "rank_math_description", "value": seo_content.get("meta_description")},
                {"key": "rank_math_focus_keyword", "value": focus_keyword},
                {"key": "rank_math_keywords", "value": all_keywords}
            ]
        }
        
        # Add tags if available
        if tags:
            data["tags"] = tags
        
        # Log the data being sent (for debugging)
        print_colored(f"Updating product {product_id} with:", Fore.BLUE)
        print_colored(f"  - Original Product Title: {product_title} (preserved)", Fore.BLUE)
        print_colored(f"  - RankMath Meta Title: {seo_content.get('meta_title')}", Fore.BLUE)
        print_colored(f"  - Focus Keyword: {focus_keyword}", Fore.BLUE)
        print_colored(f"  - Secondary Keywords: {seo_content.get('secondary_keywords')}", Fore.BLUE)
        print_colored(f"  - Tags: {seo_content.get('tags')}", Fore.BLUE)
        print_colored(f"  - Permalink: {data['slug']}", Fore.BLUE)
        
        response = requests.put(
            f"{WOOCOMMERCE_API_URL}/{product_id}",
            auth=(WOOCOMMERCE_USER, WOOCOMMERCE_PASS),
            json=data
        )
        
        if response.status_code in (200, 201):
            print_colored(f"✓ Product {product_id} updated successfully.", Fore.GREEN, True)
            # Save successful update for reference
            os.makedirs("success_logs", exist_ok=True)
            with open(f"success_logs/{product_id}_update.json", "w", encoding="utf-8") as f:
                json.dump(data, f, indent=2)
            return True
        else:
            print_colored(f"❌ Failed to update product {product_id}. Response: {response.text}", Fore.RED, True)
            failed_products.append({
                "product_name": seo_content.get("product_name"),
                "error": f"API Error: {response.status_code}",
                "details": response.text[:200] + "..." if len(response.text) > 200 else response.text
            })
            return False
    except Exception as e:
        print_colored(f"❌ Exception updating product {product_id}: {e}", Fore.RED, True)
        failed_products.append({
            "product_name": seo_content.get("product_name"),
            "error": "Exception during product update",
            "details": str(e)
        })
        return False

def update_all_product_images(product_id, focus_keyword):
    """
    Dummy function to update product images.
    Implement image update logic as needed.
    """
    print_colored(f"Updating images for product {product_id} with focus keyword '{focus_keyword}'...", Fore.BLUE)
    # Add alt text to images based on focus keyword
    # This is a placeholder - implement actual image update logic if needed
    return True

def main():
    # Print script banner
    print_colored("=" * 80, Fore.CYAN, True)
    print_colored(" WooCommerce Product SEO Optimizer with Claude SDK ", Fore.CYAN, True)
    print_colored("=" * 80, Fore.CYAN, True)
    
    # Read products from Excel
    products_df = read_products_from_excel(excel_file)
    if products_df.empty:
        print_colored("No products found to update.", Fore.RED, True)
        return
    
    total_products = len(products_df)
    successful_updates = 0
    print_colored(f"Found {total_products} products to process.", Fore.CYAN, True)
    
    for index, row in products_df.iterrows():
        # Extract product data, ensuring we use 'title' for product name
        product_title = row.get("title", f"Product_{index}")
        product_description = row.get("description", "")
        product_id = row.get("id")
        brand_name = row.get("brand", "")
        product_link = row.get("link", "")
        
        print_colored("=" * 60, Fore.BLUE)
        print_colored(f"Processing product {index+1}/{total_products}: {product_title}", Fore.BLUE, True)
        print_colored("=" * 60, Fore.BLUE)
        
        # Fetch additional info from product link if available
        additional_info = ""
        if product_link and not pd.isna(product_link):
            additional_info = fetch_additional_info(product_link)
        
        # Generate SEO content
        seo_content = generate_seo_content(product_title, product_description, additional_info, brand_name)
        if seo_content is None:
            print_colored(f"✘ Skipping product {product_title} due to SEO generation errors.", Fore.RED, True)
            continue
        
        # Update product info in WooCommerce - now passing original title
        success = update_product_info(product_id, product_title, seo_content)
        if success:
            update_all_product_images(product_id, seo_content.get('focus_keyword'))
            successful_updates += 1
            print_colored(f"✓ Updated product {product_title} (ID: {product_id}) successfully!", Fore.GREEN, True)
        else:
            print_colored(f"✘ Failed to update product {product_title} (ID: {product_id}).", Fore.RED, True)
    
    # Print summary
    print_colored("\n" + "=" * 80, Fore.CYAN, True)
    print_colored(" Summary ", Fore.CYAN, True)
    print_colored("=" * 80, Fore.CYAN, True)
    print_colored(f"Total products processed: {total_products}", Fore.CYAN)
    print_colored(f"Successfully updated: {successful_updates}", Fore.GREEN)
    print_colored(f"Failed updates: {total_products - successful_updates}", Fore.RED)
    
    # Save failed products to file
    if failed_products:
        os.makedirs("errors", exist_ok=True)
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        error_file = f"errors/failed_products_{timestamp}.json"
        with open(error_file, "w", encoding="utf-8") as f:
            json.dump(failed_products, f, indent=2)
        print_colored(f"Failed products details saved to {error_file}", Fore.YELLOW)

if __name__ == "__main__":
    main()