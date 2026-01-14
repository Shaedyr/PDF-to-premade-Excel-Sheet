# =========================
# EXCEL CLOUD TEMPLATE LOADER - DIRECT METHOD
# =========================
def load_template_from_excel_cloud():
    """
    Direct method to load Excel from Excel Cloud
    Uses browser-like headers and follows redirects
    """
    try:
        # Your Excel Cloud share link
        share_link = "https://1drv.ms/x/c/f5e2800feeb07258/IQBBPI2scMXjQ6bi18LIvXFGAWFnYqG3J_kCKfewCEid9Bc?e=ccyPnQ"
        
        st.info(f"Using Excel Cloud link: {share_link}")
        
        # Method 1: Use the share link with a direct download trick
        # OneDrive often redirects to a different URL that we need to follow
        
        # First, let's just try the link as-is with proper headers
        headers = {
            "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36",
            "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8",
            "Accept-Language": "en-US,en;q=0.5",
            "Accept-Encoding": "gzip, deflate, br",
            "DNT": "1",
            "Connection": "keep-alive",
            "Upgrade-Insecure-Requests": "1",
            "Sec-Fetch-Dest": "document",
            "Sec-Fetch-Mode": "navigate",
            "Sec-Fetch-Site": "none",
            "Sec-Fetch-User": "?1",
            "Cache-Control": "max-age=0"
        }
        
        # Create a session to handle cookies and redirects
        session = requests.Session()
        
        # First request to get redirect URL
        st.info("Step 1: Following redirects...")
        try:
            initial_response = session.head(share_link, headers=headers, timeout=30, allow_redirects=True)
            final_url = initial_response.url
            st.info(f"Redirected to: {final_url}")
        except:
            # Try with GET if HEAD fails
            initial_response = session.get(share_link, headers=headers, timeout=30, allow_redirects=True)
            final_url = initial_response.url
            st.info(f"Redirected to: {final_url}")
        
        # Method 2: Convert to direct download using known pattern
        # Check if we got a onedrive.live.com URL
        if "onedrive.live.com" in final_url:
            # This is good, try to download directly
            download_url = final_url
            
            # Replace view with download if needed
            if "redir?" in download_url:
                download_url = download_url.replace("redir?", "download?")
            elif "?" in download_url and "download=1" not in download_url:
                download_url += "&download=1"
            elif "download=1" not in download_url:
                download_url += "?download=1"
        else:
            # Try to construct direct download URL
            # Extract share token from original URL
            import base64
            import urllib.parse
            
            # The share token is the long string after /c/
            parts = share_link.split('/')
            share_token = None
            for i, part in enumerate(parts):
                if part == 'c' and i + 1 < len(parts):
                    share_token = parts[i + 1].split('?')[0]
                    break
            
            if share_token:
                # URL encode the share token
                encoded_token = urllib.parse.quote(share_token)
                download_url = f"https://api.onedrive.com/v1.0/shares/u!{encoded_token}/root/content"
                st.info(f"Constructed API URL: {download_url}")
            else:
                download_url = share_link
        
        # Method 3: Try to download
        st.info("Step 2: Downloading file...")
        response = session.get(download_url, headers=headers, timeout=30, allow_redirects=True, stream=True)
        
        st.info(f"Download response: HTTP {response.status_code}")
        st.info(f"Content-Type: {response.headers.get('content-type', 'unknown')}")
        
        if response.status_code == 200:
            # Read the content
            content = b""
            for chunk in response.iter_content(chunk_size=8192):
                if chunk:
                    content += chunk
            
            st.info(f"Downloaded {len(content)} bytes")
            
            # Check if it's a valid Excel file
            if len(content) > 1000:
                # Check for Excel signatures
                if content.startswith(b'PK') or content.startswith(b'\xD0\xCF\x11\xE0'):
                    st.success(f"✅ Successfully downloaded Excel file ({len(content)} bytes)")
                    return content
                else:
                    # Check what we actually got
                    try:
                        text_start = content[:500].decode('utf-8', errors='ignore').lower()
                        if '<html' in text_start or '<!doctype' in text_start:
                            st.error("❌ Got HTML instead of Excel. Showing first 500 chars:")
                            st.code(text_start[:500])
                            
                            # Try to find download link in HTML
                            if 'download' in text_start or '.xlsx' in text_start:
                                # Try to extract download link from HTML
                                soup = BeautifulSoup(content, 'html.parser')
                                for link in soup.find_all('a', href=True):
                                    href = link['href']
                                    if '.xlsx' in href or 'download' in href.lower():
                                        st.info(f"Found potential download link: {href}")
                                        # Try this link
                                        if href.startswith('http'):
                                            dl_response = session.get(href, headers=headers, timeout=30)
                                            if dl_response.status_code == 200:
                                                dl_content = dl_response.content
                                                if len(dl_content) > 1000 and (dl_content.startswith(b'PK') or dl_content.startswith(b'\xD0\xCF\x11\xE0')):
                                                    st.success("✅ Got Excel from HTML link!")
                                                    return dl_content
                    except:
                        pass
                    
                    st.error(f"❌ Not a valid Excel file. First 20 bytes hex: {content[:20].hex()}")
            else:
                st.error(f"❌ File too small ({len(content)} bytes)")
        else:
            st.error(f"❌ Download failed with HTTP {response.status_code}")
            
        # Method 4: Last resort - use the web viewer trick
        st.info("Step 3: Trying web viewer method...")
        web_viewer_url = "https://onedrive.live.com/download?resid=f5e2800feeb07258!107&authkey=!IQBBPI2scMXjQ6bi18LIvXFGAWFnYqG3J_kCKfewCEid9Bc"
        
        try:
            viewer_response = session.get(web_viewer_url, headers=headers, timeout=30)
            if viewer_response.status_code == 200:
                viewer_content = viewer_response.content
                if len(viewer_content) > 1000 and (viewer_content.startswith(b'PK') or viewer_content.startswith(b'\xD0\xCF\x11\xE0')):
                    st.success("✅ Web viewer method worked!")
                    return viewer_content
        except:
            pass
            
        return None
        
    except Exception as e:
        st.error(f"❌ Error loading template: {str(e)}")
        import traceback
        st.error(traceback.format_exc())
        return None
