import streamlit as st
import pandas as pd
from claude_automate_main import (
    read_products_from_excel,
    generate_seo_content,
    update_product_info,
    update_all_product_images,
    fetch_additional_info
)
import tempfile

st.title("WooCommerce Product SEO Updater")
st.markdown("Enter your credentials and upload your Excel file to update products.")

# Credentials inputs with empty default values
woocommerce_url = st.text_input("WooCommerce API URL", "")
woocommerce_user = st.text_input("WooCommerce User Key", "")
woocommerce_pass = st.text_input("WooCommerce Pass Key", "", type="password")
claude_api_key = st.text_input("Claude API Key", "", type="password")
wp_media_url = st.text_input("WordPress Media URL", "")
wp_media_user = st.text_input("WordPress Media Username", "")
wp_media_pass = st.text_input("WordPress Media Password", "", type="password")

# File uploader for the Excel file
uploaded_file = st.file_uploader("Upload Excel File", type=["xlsx"])

if st.button("Update Products"):
    # Basic validation: make sure all fields are entered
    if not (woocommerce_url and woocommerce_user and woocommerce_pass and claude_api_key and wp_media_url and wp_media_user and wp_media_pass):
        st.error("Please fill in all the credential fields.")
    elif uploaded_file is None:
        st.error("Please upload an Excel file.")
    else:
        # Save the uploaded file to a temporary file
        with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp:
            tmp.write(uploaded_file.getvalue())
            tmp_path = tmp.name

        # Read products from the temporary Excel file
        products_df = pd.read_excel(tmp_path)
        if products_df.empty:
            st.error("No products found in the uploaded file.")
        elif len(products_df) > 6:
            st.error("The uploaded Excel file must have 6 rows or less.")
        else:
            st.success(f"Found {len(products_df)} products. Starting updates...")
            successful_updates = 0
            total_products = len(products_df)
            
            # Loop through each product (assumes Excel has columns like 'id', 'title', 'description', etc.)
            for index, row in products_df.iterrows():
                product_title = row.get("title", f"Product_{index}")
                product_description = row.get("description", "")
                product_id = row.get("id")
                brand_name = row.get("brand", "")
                product_link = row.get("link", "")
                
                st.write(f"Processing **{product_title}** (ID: {product_id})...")
                
                additional_info = ""
                if pd.notna(product_link):
                    additional_info = fetch_additional_info(product_link)
                
                # Generate SEO content; note that you may need to modify your function
                seo_content = generate_seo_content(product_title, product_description, additional_info, brand_name)
                if seo_content is None:
                    st.error(f"SEO generation failed for {product_title}. Skipping...")
                    continue
                
                # Update product info using dynamic credentials
                success = update_product_info(product_id, product_title, seo_content)
                if success:
                    primary_focus_keyword = seo_content.get('primary_focus_keyword', '')
                    update_all_product_images(product_id, primary_focus_keyword)
                    st.success(f"Updated product **{product_title}** (ID: {product_id}).")
                    successful_updates += 1
                else:
                    st.error(f"Failed to update product **{product_title}** (ID: {product_id}).")
            
            st.markdown("---")
            st.info(f"Update process completed: {successful_updates} out of {total_products} products updated successfully.")
