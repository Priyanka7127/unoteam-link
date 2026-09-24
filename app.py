"""
app.py — Internal Linking Audit tool (runs on YOUR computer).

Usage:
    pip install streamlit requests beautifulsoup4 openpyxl
    streamlit run app.py

This opens a page in your browser at http://localhost:8501 with a URL
box and toggle buttons. Because it runs locally, not inside a claude.ai
page, it can make real web requests to any site you point it at.
"""
import streamlit as st
import crawler_core as cc
import report_builder as rb

st.set_page_config(page_title="UnoTeam Link Audit", page_icon="🔗", layout="centered")

st.title("🔗 UnoTeam Link Audit")
st.caption("Runs locally on your computer — makes real requests to the site you enter below.")

url = st.text_input("Website URL", placeholder="https://example.com/")

st.markdown("**Sections to audit**")
col1, col2, col3 = st.columns(3)
with col1:
    chk_content = st.checkbox("Treatments / Service Pages", value=True)
with col2:
    chk_blog = st.checkbox("Blogs", value=True)
with col3:
    chk_case = st.checkbox("Case Studies", value=True)

with st.expander("Advanced settings"):
    max_pages = st.slider("Max pages to crawl", 20, 500, 150, step=10,
                            help="Higher = more complete, but slower. Start at 150 and increase "
                                 "if your Summary tab says pages were cut off.")
    delay = st.slider("Delay between requests (seconds)", 0.0, 2.0, 0.4, step=0.1,
                        help="Be polite to the target server — don't set this to 0 on a live site "
                             "you don't own.")

run = st.button("Run Audit", type="primary", use_container_width=True)

if run:
    if not url or not url.startswith(("http://", "https://")):
        st.error("Enter a full URL starting with http:// or https://")
    elif not (chk_content or chk_blog or chk_case):
        st.error("Select at least one section to audit.")
    else:
        progress_bar = st.progress(0, text="Starting crawl...")
        status_area = st.empty()

        def progress_cb(done, total):
            pct = min(done / total, 1.0)
            progress_bar.progress(pct, text=f"Crawled {done} of up to {total} pages...")

        def log_cb(msg):
            status_area.text(msg)

        with st.spinner("Crawling — this can take a minute or two on larger sites..."):
            try:
                pages = cc.crawl_site(url, max_pages=max_pages, delay=delay,
                                        progress_cb=progress_cb, log_cb=log_cb)
            except Exception as e:
                st.error(f"Crawl failed: {e}")
                st.stop()

        progress_bar.progress(1.0, text="Crawl complete.")
        status_area.empty()

        if len(pages) <= 1:
            st.warning(
                "Only reached 1 page. This usually means the homepage didn't load, or the site "
                "blocks automated requests. Double-check the URL and try again."
            )
            st.stop()

        inbound = cc.build_inbound_index(pages)
        broken = rb.find_broken_links(pages, inbound)

        # --- Quick on-screen summary ---
        st.success(f"Crawled {len(pages)} pages.")
        counts = {}
        for cat in ("content", "blog", "case_study"):
            counts[cat] = sum(1 for d in pages.values() if d["category"] == cat and d.get("status") == 200)
        c1, c2, c3, c4 = st.columns(4)
        c1.metric("Treatment/Service pages", counts["content"])
        c2.metric("Blog posts", counts["blog"])
        c3.metric("Case studies", counts["case_study"])
        c4.metric("Broken links found", len(broken))

        orphans_content = [u for u, d in pages.items()
                            if d["category"] == "content" and d.get("status") == 200
                            and len(inbound.get(u, [])) == 0]
        if orphans_content:
            st.warning(f"{len(orphans_content)} treatment/service page(s) have zero inbound links "
                        f"(orphans) — see the Excel report for exactly which pages should link to them.")

        selected = {"content": chk_content, "blog": chk_blog, "case_study": chk_case}
        excel_bytes = rb.build_workbook(pages, selected, url)

        st.download_button(
            "⬇️ Download Excel Report",
            data=excel_bytes,
            file_name="internal_linking_audit.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
            use_container_width=True,
        )

        if len(pages) >= max_pages:
            st.info(
                f"This crawl stopped at the {max_pages}-page limit. If your site has more pages "
                "than that, raise 'Max pages to crawl' in Advanced settings and run again for a "
                "fully complete picture."
            )
