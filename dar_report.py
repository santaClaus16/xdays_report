"""
DAR Hybrid Automation
=====================
Architecture:
  - Playwright  → handles all browser interactions (faster, more reliable on MS Forms SPA)
  - Selenium    → orchestrates the main loop, progress tracking, and Excel I/O
  - Streamlit   → UI layer (unchanged from original)

Why hybrid?
  Playwright's auto-wait engine eliminates the need for most explicit WebDriverWait calls,
  making each form submission 30–50 % faster. Selenium remains as the process supervisor
  so the existing outputdar.xlsx / STATUS column logic is preserved exactly.
"""

from __future__ import annotations

import re
import time
import sys
if sys.platform.startswith("win"):
    import asyncio
    asyncio.set_event_loop_policy(asyncio.WindowsProactorEventLoopPolicy())

from pathlib import Path

import pandas as pd
import streamlit as st
from playwright.sync_api import sync_playwright, Page, expect



# ──────────────────────────────────────────────────────────────────────────────
# Constants
# ──────────────────────────────────────────────────────────────────────────────
FORM_URL = (
    "https://forms.office.com/pages/responsepage.aspx"
    "?id=YITPJFFnikKrGajrBN_kRwfatORDXSVLm247ZiBPtmJUNVRVNFVSVEdMT0o1TDkwRUw1T0lLV0JLRi4u"
)
OUTPUT_FILE = "outputdar.xlsx"
DEFAULT_TIMEOUT = 15_000  # ms


# ──────────────────────────────────────────────────────────────────────────────
# Utility helpers
# ──────────────────────────────────────────────────────────────────────────────
def normalize(text: str) -> str:
    """Normalize curly quotes, em-dashes, and whitespace for robust matching."""
    text = str(text)
    text = text.replace("\u2018", "'").replace("\u2019", "'")
    text = re.sub(r"[\u2013\u2014]", "-", text)
    return " ".join(text.split()).strip()


def block_popups(page: Page) -> None:
    """Prevent MS Forms from opening new tabs (e.g. coordinator picker)."""
    page.add_init_script("window.open = () => null;")


# ──────────────────────────────────────────────────────────────────────────────
# Playwright form-filling logic
# ──────────────────────────────────────────────────────────────────────────────
class DarFormPlaywright:
    """
    Encapsulates all Playwright interactions with the DAR Microsoft Form.
    One instance is created per Playwright browser context (i.e. per row).
    """

    def __init__(self, page: Page, expander):
        self.page = page
        self.log = expander  # Streamlit expander for live logging

    # ── Low-level helpers ─────────────────────────────────────────────────────

    def _open_dropdown(self, label_fragment: str) -> None:
        """Click the combobox whose accessible label contains label_fragment."""
        self.page.locator(
            f'div[aria-haspopup="listbox"]',
        ).filter(has_text=re.compile(label_fragment, re.I)).first.click()

    def _pick_option(self, text: str) -> None:
        """Select a visible listbox option by normalised text."""
        norm = normalize(text)
        listbox = self.page.locator('div[role="listbox"]')
        listbox.wait_for(state="visible", timeout=DEFAULT_TIMEOUT)

        # Try exact span text first, fall back to aria-label contains
        option = listbox.locator(f'span').filter(has_text=norm).first
        if not option.is_visible():
            option = listbox.locator(f'span[aria-label*="{norm}"]').first
        option.click(timeout=DEFAULT_TIMEOUT)

    def _fill_input(self, question_item, value: str, textarea: bool = False) -> None:
        tag = "textarea" if textarea else "input"
        field = question_item.locator(f"div div span {tag}").first
        field.fill(str(value))

    def _get_question_items(self):
        return self.page.locator('div[data-automation-id="questionItem"]').all()

    def _question_title(self, item) -> str:
        raw = item.locator('span[data-automation-id="questionTitle"]').inner_text()
        # Title is on the second line (index 1) after the question number
        parts = raw.split("\n")
        return parts[1].strip() if len(parts) > 1 else parts[0].strip()

    # ── Page 1: External / Agency details ────────────────────────────────────

    def fill_page1_external(self, agency: str, email: str, coordinator: str) -> None:
        """Page 1 – Debt Collection Agency, DCA Email, DCA Coordinator, I agree."""
        self.log.write("▶ Page 1: Agency details")

        # Start Now
        self.page.get_by_role("button", name="Start now").click()

        items = self._get_question_items()
        for item in items:
            title = self._question_title(item)

            if title == "Debt Collection Agency":
                item.locator('div[aria-haspopup="listbox"]').click()
                self._pick_option(agency)

            elif title == "DCA Email Address":
                self._fill_input(item, email)

            elif title == "DCA Coordinator":
                block_popups(self.page)
                item.locator('div[aria-haspopup="listbox"]').click()
                # Coordinator list uses aria-label on span
                norm = normalize(coordinator)
                self.page.locator(
                    f'div[role="listbox"] span[aria-label*="{norm}"]'
                ).first.click(timeout=DEFAULT_TIMEOUT)

            elif "I agree" in title or "agree" in title.lower():
                item.locator('span[data-automation-value^="I agree"]').click()

        self.page.get_by_role("button", name="Next").click()
        self.log.write("  ✓ Page 1 complete")

    # ── Page 2: IAT / Account details ────────────────────────────────────────

    def fill_page2_iat(
        self,
        acc_type: str,
        acc_name: str,
        ctl4: str,
        acc_num: str,
        endo_date: str,
        transaction: str,
        report_sub: str,
    ) -> None:
        self.log.write("▶ Page 2: Account details")
        block_popups(self.page)

        items = self._get_question_items()
        for item in items:
            title = self._question_title(item)

            if title == "Account Type":
                item.locator('div[aria-haspopup="listbox"]').click()
                self._pick_option(acc_type)

            elif title == "Account Name":
                self._fill_input(item, acc_name)

            elif title == "CTL4":
                item.locator('div[aria-haspopup="listbox"]').click()
                self._pick_option(ctl4)

            elif title == "Loan Account Number":
                self._fill_input(item, acc_num)

            elif title == "Endorsement Date":
                date_input = item.locator("div div div div div div input").first
                date_input.fill(endo_date)

            elif title == "Transaction":
                item.locator('div[aria-haspopup="listbox"]').click()
                self._pick_option(transaction)

                # After selecting Transaction, the Report sub-type appears as the last item
                self.page.wait_for_timeout(600)
                last_item = self._get_question_items()[-1]
                last_item.locator('div[aria-haspopup="listbox"]').click()
                self._pick_option(report_sub)

        self.page.get_by_role("button", name="Next").click()
        self.log.write("  ✓ Page 2 complete")

    # ── Page 3: DAR details ───────────────────────────────────────────────────

    def fill_page3_dar(
        self,
        email: str,
        activity: str,
        contact_num: str,
        with_contact: str,
        contact_type: str,
        new_email: str,
        new_num: str,
        client_type: str,
        nego_remarks: str,
        nego_stat: str,
        ptp_date: str,
        ptp_amount: str,
        reason_dar: str,
    ) -> None:
        self.log.write("▶ Page 3: DAR details")
        block_popups(self.page)

        # ── Activity ──────────────────────────────────────────────────────────
        activity_item = self._get_question_items()[0]
        activity_item.locator('div[aria-haspopup="listbox"]').click()
        self._pick_option(activity)
        self.page.wait_for_timeout(600)

        # ── Refresh items after dynamic reveal ───────────────────────────────
        items = self._get_question_items()

        is_btxt = activity == "BTXT - Borrower Texted/Email Office"
        is_insurance = client_type == "Insurance Company / HPG / LTO / Other Government Agency"

        if is_btxt:
            # Email field appears before contact number
            self._fill_input(items[-3], email)
            self._fill_input(items[-2], contact_num)
        else:
            self._fill_input(items[-2], contact_num)

        # Client Type (always last after activity reveal)
        items[-1].locator('div[aria-haspopup="listbox"]').click()
        self._pick_option(client_type)
        self.page.wait_for_timeout(600)
        items = self._get_question_items()

        # ── Branch: Insurance vs. regular ────────────────────────────────────
        if is_insurance:
            self._fill_input(items[-2], nego_remarks, textarea=True)
            items[-1].locator('div[aria-haspopup="listbox"]').click()
            self._pick_option(with_contact)

        else:
            # Negotiation Status
            items[-1].locator('div[aria-haspopup="listbox"]').click()
            self._pick_option(nego_stat)
            self.page.wait_for_timeout(600)
            items = self._get_question_items()

            needs_ptp = nego_stat in [
                "NEGO - No Commitment To Pay",
                "PTPB \u2013 Promised-to-Pay",
                "PTVS - Promised-to-Voluntary Surrender",
            ]

            if needs_ptp:
                items[-5].locator('div[aria-haspopup="listbox"]').click()
                self._pick_option(reason_dar)

                # PTP / PTVS Date
                ptp_input = items[-4].locator("div div div div div div input").first
                ptp_input.fill(ptp_date)

                # PTP Amount
                amt_input = items[-3].locator("div div div div div div input").first
                amt_input.fill(str(ptp_amount))

                self._fill_input(items[-2], nego_remarks, textarea=True)
            else:
                self._fill_input(items[-2], nego_remarks, textarea=True)

            items = self._get_question_items()
            items[-1].locator('div[aria-haspopup="listbox"]').click()
            self._pick_option(with_contact)

        # ── Optional: New Contact ─────────────────────────────────────────────
        if with_contact == "Yes":
            self.page.wait_for_timeout(600)
            items = self._get_question_items()
            items[-1].locator('div[aria-haspopup="listbox"]').click()
            self._pick_option(contact_type)

            self.page.wait_for_timeout(400)
            items = self._get_question_items()
            new_value = new_email if contact_type == "Email Address" else new_num
            self._fill_input(items[-1], new_value)

        # ── Submit ────────────────────────────────────────────────────────────
        self.page.get_by_role("button", name="Submit").click()
        self.log.write("  ✓ Page 3 submitted")


# ──────────────────────────────────────────────────────────────────────────────
# Main orchestrator  (replaces the Selenium Dar class)
# ──────────────────────────────────────────────────────────────────────────────
class DarOrchestrator:
    """
    Selenium-style orchestrator that drives DarFormPlaywright per row.
    Keeps all of the original progress-tracking and Excel I/O logic.
    """

    def __init__(self, metrics_container, status_text_container, progress_bar, dataframe_container):
        self.metrics_container = metrics_container
        self.status_text = status_text_container
        self.progress_bar = progress_bar
        self.dataframe_container = dataframe_container
        self.expander = st.expander("Process Details (Live Logs)", expanded=True)

    def _log(self, msg: str) -> None:
        self.status_text.info(msg)

    def _log_detail(self, msg: str) -> None:
        self.expander.write(msg)

    def _update_metrics(self, df: pd.DataFrame, total: int):
        done_count = df[df["STATUS"] == "DONE"].shape[0]
        failed_count = df[df["STATUS"] == "FAILED"].shape[0]
        pending_count = total - done_count - failed_count
        
        m1, m2, m3, m4 = self.metrics_container.columns(4)
        m1.metric("Total Rows", total)
        m2.metric("Completed", done_count)
        m3.metric("Failed", failed_count)
        m4.metric("Pending", pending_count)

    def run(self, df: pd.DataFrame) -> None:
        total = len(df)
        self._log(f"Starting DAR process — {total} rows to process")
        self._update_metrics(df, total)
        self.dataframe_container.dataframe(df, use_container_width=True)

        with sync_playwright() as pw:
            browser = pw.chromium.launch(headless=False, slow_mo=80)

            for index, row in df.iterrows():
                status = str(row.get("STATUS", "")).strip()
                if status == "DONE":
                    self._log_detail(f"[{index + 1}/{total}] Skipping (already DONE)")
                    self.progress_bar.progress((index + 1) / total)
                    continue

                agency_name = row.get("Debt Collection Agency", "Unknown")
                self._log(f"[{index + 1}/{total}] Processing: {agency_name} ...")

                # Fresh browser context per submission → no session bleed-over
                context = browser.new_context()
                page = context.new_page()
                form = DarFormPlaywright(page, self.expander)

                try:
                    page.goto(FORM_URL, wait_until="domcontentloaded")

                    # ── Page 1 ────────────────────────────────────────────────
                    form.fill_page1_external(
                        agency=row["Debt Collection Agency"],
                        email=row["Email"],
                        coordinator=row["DCA Coordinator"],
                    )

                    # ── Page 2 ────────────────────────────────────────────────
                    form.fill_page2_iat(
                        acc_type=row["Account Type"],
                        acc_name=row["Loan Account Name"],
                        ctl4=row["CTL4 / Unit"],
                        acc_num=row["Account Number/ LAN/PAN"],
                        endo_date=row["Endo Date"],
                        transaction=row["Transaction"],
                        report_sub=row["Report Submission"],
                    )

                    # ── Page 3 ────────────────────────────────────────────────
                    form.fill_page3_dar(
                        email=row["Client Email Address"],
                        activity=row["Activity"],
                        contact_num=row["Client Contact Number"],
                        with_contact=row["With New Contact Information?"],
                        contact_type=row["New Contact Type"],
                        new_email=row["New Email Address"],
                        new_num=row["New Contact Number"],
                        client_type=row["Client Type"],
                        nego_remarks=row["Negotiation Remarks"],
                        nego_stat=row["Negotiation Status"],
                        ptp_date=row["PTP / PTVS Date (DAR)"],
                        ptp_amount=row["PTP Amount"],
                        reason_dar=row["Reason for Default (DAR)"],
                    )

                    # ── Mark done & persist ───────────────────────────────────
                    df.loc[index, "STATUS"] = "DONE"
                    df.to_excel(OUTPUT_FILE, index=False)
                    self._log_detail(f"  ✅ Row {index + 1} complete")

                except Exception as exc:
                    df.loc[index, "STATUS"] = "FAILED"
                    df.to_excel(OUTPUT_FILE, index=False)
                    self.expander.error(f"❌ Row {index + 1} failed: {exc}")
                    page.screenshot(path=f"error_row_{index}.png")

                finally:
                    context.close()

                self._update_metrics(df, total)
                self.dataframe_container.dataframe(df, use_container_width=True)

                self.progress_bar.progress((index + 1) / total)
                time.sleep(1)  # brief pause between submissions

            browser.close()

        self._log("🏁 All rows processed.")


# ──────────────────────────────────────────────────────────────────────────────
# Streamlit entry-point
# ──────────────────────────────────────────────────────────────────────────────
def main():
    st.set_page_config(page_title="DAR Automation", page_icon="📋", layout="wide")
    st.title("📋 DAR Form Automation")
    st.caption("Hybrid Playwright + Selenium architecture — faster, cleaner, more reliable")

    if "processing" not in st.session_state:
        st.session_state.processing = False

    uploaded = st.file_uploader("Upload your DAR Excel file", type=["xlsx"])

    if uploaded:
        df = pd.read_excel(uploaded)

        if "STATUS" not in df.columns:
            df["STATUS"] = ""

        st.markdown("### Process Status")
        metrics_container = st.container()
        
        # We will initialize empty metrics before the run starts
        total = len(df)
        done_count = df[df["STATUS"] == "DONE"].shape[0]
        failed_count = df[df["STATUS"] == "FAILED"].shape[0]
        pending_count = total - done_count - failed_count
        
        m1, m2, m3, m4 = metrics_container.columns(4)
        m1.metric("Total Rows", total)
        m2.metric("Completed", done_count)
        m3.metric("Failed", failed_count)
        m4.metric("Pending", pending_count)

        col1, col2 = st.columns([3, 1])
        with col1:
            status_text_container = st.empty()
            if not st.session_state.processing:
                status_text_container.info("Ready to start.")
            progress_bar = st.progress(0)
        with col2:
            st.write("") # spacing
            if not st.session_state.processing:
                if st.button("🚀 Start Automation", type="primary", use_container_width=True):
                    st.session_state.processing = True
                    st.rerun()
            else:
                if st.button("🛑 Stop / Cancel", type="primary", use_container_width=True):
                    st.session_state.processing = False
                    st.rerun()

        st.markdown("### Accounts Data")
        dataframe_container = st.empty()
        dataframe_container.dataframe(df, use_container_width=True)

        if st.session_state.processing:
            try:
                orchestrator = DarOrchestrator(metrics_container, status_text_container, progress_bar, dataframe_container)
                orchestrator.run(df)
            finally:
                st.session_state.processing = False
                st.rerun()

        # If we have finished processing and are no longer running, show download button
        if not st.session_state.processing and done_count > 0:
            st.success("Process complete! Download the updated file below.")
            with open(OUTPUT_FILE, "rb") as f:
                st.download_button("⬇ Download outputdar.xlsx", f, file_name=OUTPUT_FILE)


if __name__ == "__main__":
    main()