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
            text = response.text[:5000]
            print_colored(f"Successfully fetched {len(text)} characters of additional info", Fore.GREEN)
            return text
        else:
            print_colored(f"Failed to fetch info: Status code {response.status_code}", Fore.RED)
            return f"Failed to fetch info: Status code {response.status_code}"
    except Exception as e:
        print_colored(f"Error fetching additional info: {e}", Fore.RED)
        return f"Error fetching additional info: {e}"

def create_optimized_permalink(product_name, focus_keyword, max_length=60):
    """
    Create an SEO-friendly permalink that starts with the focus keyword.
    Ensures the permalink is no longer than max_length characters.
    """
    simplified_keyword = focus_keyword.lower()
    simplified_keyword = re.sub(r'[^a-z0-9\s]', '', simplified_keyword)
    simplified_keyword = re.sub(r'\s+', '-', simplified_keyword.strip())
    
    base_permalink = simplified_keyword
    
    if len(base_permalink) < max_length - 5:
        simplified_name = re.sub(r'[^a-z0-9\s]', '', product_name.lower())
        simplified_name = re.sub(r'\s+', '-', simplified_name.strip())
        
        remaining_space = max_length - len(base_permalink) - 1
        if simplified_name and simplified_name != simplified_keyword and remaining_space > 5:
            name_part = simplified_name[:remaining_space]
            if '-' in name_part:
                name_part = name_part.rsplit('-', 1)[0]
            base_permalink += f"-{name_part}"
    
    if len(base_permalink) > max_length:
        if '-' in base_permalink[:max_length]:
            base_permalink = base_permalink[:max_length].rsplit('-', 1)[0]
        else:
            base_permalink = base_permalink[:max_length]
        
    base_permalink = base_permalink.rstrip('-')
    
    return base_permalink

def validate_seo_content(product_name, sections):
    """Validate that all required SEO content sections are present."""
    required_fields = ['LONG DESCRIPTION', 'SHORT DESCRIPTION', 'META TITLE', 'META DESCRIPTION', 'FOCUS KEYWORDS', 'SECONDARY KEYWORDS', 'TAGS', 'PERMALINK']
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
      - Permalink <= 60 chars (5 pts)
      - Meta description length between 140-155 (10 pts)
      - Meta description contains focus keyword (10 pts)
    """
    score = 0
    focus_keywords = seo_content.get("focus_keywords", "").lower()
    focus_keywords_list = [kw.strip() for kw in focus_keywords.split(',')]
    primary_focus = focus_keywords_list[0] if focus_keywords_list else ""
    
    meta_title = seo_content.get("meta_title", "")
    meta_desc = seo_content.get("meta_description", "")
    permalink = seo_content.get("permalink", "").lower()
    power_words = ["exclusive", "premium", "luxurious", "ultimate", "authentic", "stunning"]

    if meta_title.lower().startswith(primary_focus):
        score += 25
        print_colored("✓ Meta title starts with primary focus keyword", Fore.GREEN)
    else:
        print_colored("✗ Meta title doesn't start with primary focus keyword", Fore.RED)

    if any(pw in meta_title.lower() for pw in power_words):
        score += 15
        print_colored("✓ Meta title contains a power word", Fore.GREEN)
    else:
        print_colored("✗ Meta title doesn't contain a power word", Fore.RED)

    if re.search(r'\d', meta_title):
        score += 10
        print_colored("✓ Meta title contains a number", Fore.GREEN)
    else:
        print_colored("✗ Meta title doesn't contain a number", Fore.RED)

    if len(meta_title) <= 60:
        score += 10
        print_colored(f"✓ Meta title length is good: {len(meta_title)}/60", Fore.GREEN)
    else:
        print_colored(f"✗ Meta title too long: {len(meta_title)}/60", Fore.RED)

    focus_in_permalink = primary_focus.replace(' ', '-')
    if permalink.startswith(focus_in_permalink):
        score += 15
        print_colored("✓ Permalink starts with focus keyword", Fore.GREEN)
    else:
        print_colored("✗ Permalink doesn't start with focus keyword", Fore.RED)

    if len(permalink) <= 60:
        score += 5
        print_colored(f"✓ Permalink length is good: {len(permalink)}/60", Fore.GREEN)
    else:
        print_colored(f"✗ Permalink too long: {len(permalink)}/60", Fore.RED)

    if 140 <= len(meta_desc) <= 155:
        score += 10
        print_colored(f"✓ Meta description length is ideal: {len(meta_desc)}", Fore.GREEN)
    else:
        print_colored(f"✗ Meta description length is not ideal: {len(meta_desc)}", Fore.RED)

    if primary_focus in meta_desc.lower():
        score += 10
        print_colored("✓ Meta description contains focus keyword", Fore.GREEN)
    else:
        print_colored("✗ Meta description doesn't contain focus keyword", Fore.RED)

    return score

def create_seo_title(focus_keyword, product_name, max_length=60):
    """
    Create an SEO-optimized title that always begins with the focus keyword.
    Includes a power word and a number to boost SEO score.
    Ensures no colon at the end of the title.
    """
    focus_keyword = focus_keyword.strip()
    focus_capitalized = ' '.join(word.capitalize() for word in focus_keyword.split())
    
    power_words = ["Exclusive", "Premium", "Luxury", "Authentic", "Ultimate", "Elegant"]
    sentiment_words = ["Experience", "Collection", "Selection", "Choice", "Quality", "Perfection"]
    
    power_word = random.choice(power_words)
    sentiment_word = random.choice(sentiment_words)
    number = random.randint(1, 99)
    
    seo_title = f"{focus_capitalized} - {power_word} {number}% {sentiment_word}"
    
    if len(seo_title) > max_length:
        seo_title = f"{focus_capitalized} - {power_word} {number}"
    
    if len(seo_title) > max_length:
        seo_title = focus_capitalized
        
    if len(seo_title) > max_length:
        seo_title = seo_title[:max_length]
    
    seo_title = seo_title.rstrip(':')
    
    return seo_title

def improve_seo_fields(seo_content, product_name):
    """
    Improve SEO fields to ensure they meet all RankMath criteria,
    particularly ensuring the focus keyword appears at the beginning of the meta title.
    """
    focus_keywords = seo_content.get("focus_keywords", "").strip()
    focus_keywords_list = [kw.strip() for kw in focus_keywords.split(',')]
    
    if not focus_keywords_list or len(focus_keywords_list) < 3:
        words = product_name.split()
        primary_focus = focus_keywords_list[0] if focus_keywords_list else words[0] if words else product_name
        
        while len(focus_keywords_list) < 3:
            if len(words) > len(focus_keywords_list):
                new_keyword = words[len(focus_keywords_list)]
                focus_keywords_list.append(new_keyword)
            else:
                modifiers = ["best", "quality", "premium"]
                new_keyword = f"{modifiers[len(focus_keywords_list) % 3]} {primary_focus}"
                focus_keywords_list.append(new_keyword)
        
        focus_keywords = ", ".join(focus_keywords_list)
        seo_content["focus_keywords"] = focus_keywords
    
    primary_focus = focus_keywords_list[0]
    
    meta_title = seo_content.get("meta_title", "")
    # Generate a new SEO title for meta title if necessary.
    if not meta_title.lower().startswith(primary_focus.lower()) or meta_title.endswith(':'):
        print_colored("⚠️ Rebuilding meta title to start with focus keyword and remove ending colon", Fore.YELLOW)
        meta_title = create_seo_title(primary_focus, product_name)
    
    permalink = seo_content.get("permalink", "")
    focus_in_permalink = primary_focus.lower().replace(' ', '-')
    if not permalink.startswith(focus_in_permalink) or len(permalink) > 60:
        print_colored("⚠️ Rebuilding permalink to start with focus keyword and be under 60 chars", Fore.YELLOW)
        permalink = create_optimized_permalink(product_name, primary_focus)
    
    meta_desc = seo_content.get("meta_description", "")
    if primary_focus.lower() not in meta_desc.lower():
        print_colored("⚠️ Rebuilding meta description to include focus keyword", Fore.YELLOW)
        meta_desc = f"{primary_focus} offers an exclusive, luxurious experience. Shop now for the ultimate choice that will delight you every day!"
    
    if len(meta_desc) < 140:
        extra_text = " Perfect for all occasions. Try it today and experience the difference!"
        meta_desc += extra_text[:140 - len(meta_desc)]
    elif len(meta_desc) > 155:
        meta_desc = meta_desc[:152] + "..."
    
    seo_content["meta_title"] = meta_title
    seo_content["permalink"] = permalink
    seo_content["meta_description"] = meta_desc
    seo_content["focus_keywords"] = focus_keywords
    seo_content["focus_keywords_list"] = focus_keywords_list
    seo_content["primary_focus_keyword"] = primary_focus
    
    return seo_content

def generate_seo_content(product_name, product_description, additional_info="", brand_name=""):
    """
    Generate SEO content using the Claude API based on provided product data.
    Uses the official Anthropic Python client.
    """
    system_prompt = (
        "You are an expert eCommerce SEO product description writer specializing in optimizing product content. "
        "Your task is to write detailed and SEO-optimized product descriptions based on the provided information.\n\n"
        "Focus on creating content that ranks well in RankMath plugin. Critical requirements:\n"
        "- SEO Title MUST start with the Primary Focus Keyword exactly and MUST NOT end with a colon\n"
        "- Permalink MUST start with the Primary Focus Keyword and MUST be under 60 characters\n"
        "- Content should be clean HTML without Markdown formatting\n\n"
        "Content Requirements:\n"
        "1. Long Description (300+ words, HTML format):\n"
        "   - Include detailed and informative content optimized for SEO\n"
        "   - Use <strong> tags for highlighting important keywords (not Markdown)\n"
        "   - Start with the Primary Focus Keyword and repeat it naturally\n"
        "   - Include the Focus Keywords in subheadings (<h2>, <h3>, <h4>)\n"
        "   - Include a Product Information Table (Size, Color, Material, Brand Name)\n"
        "   - Include Key Features, Benefits, and overview\n"
        "   - Answer one frequently searched question related to the product\n"
        "   - Use emoticons/icons to evoke emotional connection\n"
        "   - Include 3-4 internal links to related products\n\n"
        "   - Include enternal links to related categories just use https://xsellpoint.com/product-category/fragrance/gender-international/ and https://xsellpoint.com/product-category/fragrance/gender-international/ and https://xsellpoint.com/product-category/makeup integrate the links normally in the text with clickable text\n\n "
        "2. Short Description (50 words max):\n"
        "   - Concise and engaging, highlighting uniqueness and key features\n"
        "   - Provided as plain text without any Markdown formatting\n\n"
        "3. SEO Elements (Optimized for Rank Math SEO Plugin):\n"
        "   - SEO Meta Title: MUST start with the exact Primary Focus Keyword, be under 60 characters, include a power word and a number, and MUST NOT end with a colon\n"
        "   - SEO Permalink: MUST start with the Primary Focus Keyword and be URL-friendly, MAXIMUM 60 CHARACTERS\n"
        "   - Meta Description: 140-155 characters, must include the Primary Focus Keyword, with a call to action\n"
        "   - Focus Keywords: Generate EXACTLY THREE focus keywords (primary, secondary, and tertiary) separated by commas\n"
        "   - Secondary Keywords: Generate EXACTLY TWO secondary keywords that complement the focus keywords\n"
        "   - Tags: Generate EXACTLY THREE product tags that are relevant to the product\n\n"
        "Output MUST include these EXACT section headers in your response:\n"
        "LONG DESCRIPTION:\n"
        "SHORT DESCRIPTION:\n"
        "META TITLE:\n"
        "META DESCRIPTION:\n"
        "FOCUS KEYWORDS:\n"
        "SECONDARY KEYWORDS:\n"
        "TAGS:\n"
        "PERMALINK:\n"
        "Do not include any Markdown formatting like ``` or ** in your output."
    )
    
    user_message = f"""Product Name: {product_name}
Product Description: {product_description if product_description else 'Not available'}
Brand: {brand_name if brand_name else 'Not specified'}
Additional Information: {additional_info if additional_info else 'None available'}

Generate the comprehensive SEO content following the EXACT format specified. Remember:
1. The META TITLE MUST start with the Primary Focus Keyword exactly and MUST NOT end with a colon
2. The PERMALINK MUST start with the Primary Focus Keyword and be MAXIMUM 60 CHARACTERS
3. Generate EXACTLY THREE focus keywords separated by commas
4. Generate EXACTLY TWO secondary keywords
5. Generate EXACTLY THREE product tags
"""
    try:
        client = Anthropic(api_key=CLAUDE_API_KEY)
        
        print_colored("🚀 Sending request to Claude API using official Python SDK...", Fore.BLUE)
        start_time = datetime.now()
        
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
        
        generated_text = response.content[0].text
        print_colored("✅ Received content from Claude API", Fore.GREEN)
        
        os.makedirs("debug", exist_ok=True)
        safe_name = re.sub(r'[^\w\-_\. ]', '_', product_name)
        with open(f"debug/{safe_name}_raw_response.txt", "w", encoding="utf-8") as f:
            f.write(generated_text)
        
        preview = "\n".join(generated_text.split('\n')[:10])
        print_colored("📄 Content preview (first few lines):", Fore.CYAN)
        print(preview)
        
        sections = {
            'LONG DESCRIPTION': '',
            'SHORT DESCRIPTION': '',
            'META TITLE': '',
            'META DESCRIPTION': '',
            'FOCUS KEYWORDS': '',
            'SECONDARY KEYWORDS': '',
            'TAGS': '',
            'PERMALINK': ''
        }
        current_section = None
        content_buffer = []
        
        for line in generated_text.split('\n'):
            line = line.strip()
            if not line:
                continue
                
            is_section_header = False
            for section in sections:
                header = section + ":"
                if line.upper().startswith(header) or line.upper() == section:
                    if current_section and content_buffer:
                        sections[current_section] = "\n".join(content_buffer).strip()
                    current_section = section
                    content_buffer = []
                    is_section_header = True
                    break
            
            if not is_section_header and current_section:
                content_buffer.append(line)
        
        if current_section and content_buffer:
            sections[current_section] = "\n".join(content_buffer).strip()
        
        if not validate_seo_content(product_name, sections):
            print_colored("⚠️ Some required SEO fields are missing. Attempting to auto-correct...", Fore.YELLOW)
        
        focus_keywords = sections['FOCUS KEYWORDS'].strip()
        secondary_keywords = sections['SECONDARY KEYWORDS'].strip()
        long_desc = clean_html_content(sections['LONG DESCRIPTION'])
        short_desc = clean_html_content(sections['SHORT DESCRIPTION'])
        meta_title = sections['META TITLE'].strip()
        meta_description = sections['META DESCRIPTION'].strip()
        permalink = sections['PERMALINK'].strip()
        tags = sections['TAGS'].strip()
        
        all_keywords = f"{focus_keywords}, {secondary_keywords}"
        
        seo_content = {
            "long_description": long_desc,
            "short_description": short_desc,
            "meta_title": meta_title,
            "meta_description": meta_description,
            "focus_keywords": focus_keywords,
            "secondary_keywords": secondary_keywords,
            "all_keywords": all_keywords,
            "tags": tags,
            "permalink": permalink,
            "product_name": product_name
        }
        
        seo_content = improve_seo_fields(seo_content, product_name)
        
        print_colored("SEO Score Breakdown:", Fore.MAGENTA, True)
        final_score = calculate_seo_score(seo_content)
        print_colored(f"Final SEO score: {final_score}/100", Fore.MAGENTA, True)
        
        internal_links = (
            '<p>Explore related categories: '
            '<a href="https://xsellpoint.com/product-category/fragrance/">Fragrance</a> | '
            '<a href="https://xsellpoint.com/product-category/fragrance/gender-international/">Gender International Fragrance</a> | '
            '<a href="https://xsellpoint.com/product-category/makeup">Makeup</a>'
            '</p>'
        )
        seo_content["long_description"] += "\n" + internal_links
        
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
    Keeps the original product title as provided in the Excel sheet.
    """
    try:
        focus_keywords = seo_content.get("focus_keywords", "")
        all_keywords = seo_content.get("all_keywords", "")
        
        tags = []
        if seo_content.get("tags"):
            tag_list = [tag.strip() for tag in seo_content.get("tags").split(",") if tag.strip()]
            if tag_list:
                tags = [{"name": tag} for tag in tag_list]
        
        data = {
            "description": seo_content.get("long_description"),
            "short_description": seo_content.get("short_description"),
            "slug": seo_content.get("permalink"),
            "meta_data": [
                {"key": "rank_math_title", "value": seo_content.get("meta_title")},
                {"key": "rank_math_description", "value": seo_content.get("meta_description")},
                {"key": "rank_math_focus_keyword", "value": focus_keywords},
                {"key": "rank_math_keywords", "value": all_keywords}
            ]
        }
        
        if tags:
            data["tags"] = tags
        
        print_colored(f"Updating product {product_id} with:", Fore.BLUE)
        print_colored(f"  - Original Product Title: {product_title} (preserved)", Fore.BLUE)
        print_colored(f"  - RankMath Meta Title: {seo_content.get('meta_title')}", Fore.BLUE)
        print_colored(f"  - Focus Keywords: {focus_keywords}", Fore.BLUE)
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
    return True

def main():
    print_colored("=" * 80, Fore.CYAN, True)
    print_colored(" WooCommerce Product SEO Optimizer with Claude SDK ", Fore.CYAN, True)
    print_colored("=" * 80, Fore.CYAN, True)
    
    products_df = read_products_from_excel(excel_file)
    if products_df.empty:
        print_colored("No products found to update.", Fore.RED, True)
        return
    
    total_products = len(products_df)
    successful_updates = 0
    print_colored(f"Found {total_products} products to process.", Fore.CYAN, True)
    
    for index, row in products_df.iterrows():
        product_title = row.get("title", f"Product_{index}")
        product_description = row.get("description", "")
        product_id = row.get("id")
        brand_name = row.get("brand", "")
        product_link = row.get("link", "")
        
        print_colored("=" * 60, Fore.BLUE)
        print_colored(f"Processing product {index+1}/{total_products}: {product_title}", Fore.BLUE, True)
        print_colored("=" * 60, Fore.BLUE)
        
        additional_info = ""
        if product_link and not pd.isna(product_link):
            additional_info = fetch_additional_info(product_link)
        
        seo_content = generate_seo_content(product_title, product_description, additional_info, brand_name)
        if seo_content is None:
            print_colored(f"✘ Skipping product {product_title} due to SEO generation errors.", Fore.RED, True)
            continue
        
        success = update_product_info(product_id, product_title, seo_content)
        if success:
            primary_focus_keyword = seo_content.get('primary_focus_keyword')
            update_all_product_images(product_id, primary_focus_keyword)
            successful_updates += 1
            print_colored(f"✓ Updated product {product_title} (ID: {product_id}) successfully!", Fore.GREEN, True)
        else:
            print_colored(f"✘ Failed to update product {product_title} (ID: {product_id}).", Fore.RED, True)
    
    print_colored("\n" + "=" * 80, Fore.CYAN, True)
    print_colored(" Summary ", Fore.CYAN, True)
    print_colored("=" * 80, Fore.CYAN, True)
    print_colored(f"Total products processed: {total_products}", Fore.CYAN)
    print_colored(f"Successfully updated: {successful_updates}", Fore.GREEN)
    print_colored(f"Failed updates: {total_products - successful_updates}", Fore.RED)
    
    if failed_products:
        os.makedirs("errors", exist_ok=True)
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        error_file = f"errors/failed_products_{timestamp}.json"
        with open(error_file, "w", encoding="utf-8") as f:
            json.dump(failed_products, f, indent=2)
        print_colored(f"Failed products details saved to {error_file}", Fore.YELLOW)

if __name__ == "__main__":
    main()
