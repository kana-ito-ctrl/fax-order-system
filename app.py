"""FAX発注書 自動処理システム - Streamlit App"""
import streamlit as st
import pandas as pd
import json
import os
from io import BytesIO
from pdf_generator import gen_sylvia_pdf, gen_haruna_pdf, load_product_master, load_ddc_master, load_staff
from ocr_module import pdf_to_images, ocr_fax_page, match_products, match_ddc, match_ddc_candidates, match_product_candidates, save_new_ddc

# ========== Page Config ==========
st.set_page_config(
    page_title="FAX発注書 自動処理システム",
    page_icon="📠",
    layout="wide"
)

# ========== Load Masters ==========
@st.cache_data
def get_product_master():
    return load_product_master()

@st.cache_data
def get_ddc_master():
    return load_ddc_master()

@st.cache_data
def get_staff():
    return load_staff()

pm = get_product_master()
ddc = get_ddc_master()
staff_data = get_staff()
staff_names = [s["name"] for s in staff_data["staff"]]

# ========== Header ==========
st.title("📠 FAX発注書 自動処理システム")
st.caption("株式会社TWO 事業管理部")

# ========== Tabs ==========
tab_auto, tab_manual = st.tabs(["📄 FAX読取モード", "✏️ 手動入力モード"])

# ============================================================
# FAX読取モード
# ============================================================
with tab_auto:
    st.header("STEP 1：FAX PDFアップロード")

    uploaded = st.file_uploader(
        "FAXで受信したPDFをアップロード",
        type=["pdf"],
        accept_multiple_files=True,
        key="fax_upload"
    )

    if uploaded:
        for file in uploaded:
            st.subheader(f"📎 {file.name}")
            pdf_bytes = file.read()

            # ページ分割
            with st.spinner("PDFをページ分割中..."):
                pages = pdf_to_images(pdf_bytes)
            st.success(f"{len(pages)}ページ検出")

            for page_info in pages:
                with st.expander(f"📄 ページ {page_info['page']}", expanded=True):
                    col_img, col_result = st.columns([1, 1])

                    with col_img:
                        import base64
                        img_data = base64.b64decode(page_info["base64"])
                        st.image(img_data, caption=f"ページ {page_info['page']}", use_container_width=True)

                    with col_result:
                        # STEP 2: OCR
                        if st.button(f"🔍 読み取り開始", key=f"ocr_{file.name}_{page_info['page']}"):
                            with st.spinner("Claude Vision APIで読み取り中..."):
                                ocr_result = ocr_fax_page(page_info["base64"])

                            if "error" in ocr_result:
                                st.error(ocr_result["error"])
                            else:
                                data_key = f"data_ocr_{file.name}_{page_info['page']}"
                                st.session_state[data_key] = ocr_result
                                st.success("読み取り完了！")

                        # STEP 3: 確認・マッチング
                        ocr_key = f"data_ocr_{file.name}_{page_info['page']}"
                        if ocr_key in st.session_state:
                            ocr_result = st.session_state[ocr_key]
                            st.markdown("---")
                            st.subheader("STEP 3：確認・マッチング")

                            # 基本情報
                            col1, col2, col3 = st.columns(3)
                            order_no = col1.text_input("オーダーNO", ocr_result.get("order_no", ""), key=f"on_{ocr_key}")
                            delivery_date = col2.text_input("納品日", ocr_result.get("delivery_date", ""), key=f"dd_{ocr_key}")
                            from datetime import date
                            order_date = col3.text_input("発注日", str(date.today()), key=f"od_{ocr_key}")

                            dest_name = st.text_input("納品先", ocr_result.get("delivery_dest", ""), key=f"dn_{ocr_key}")

                            # DDCマッチング（完全一致 → 候補選択 → 手動入力）
                            # ============================================================
                            ddc_result = match_ddc_candidates(dest_name, ddc)
                            ddc_match_key = f"ddc_selected_{ocr_key}"

                            if ddc_result["exact_match"]:
                                # --- 完全一致：自動マッチング ---
                                ddc_match = ddc_result["matched_row"]
                                ddc_match["matched"] = True
                                st.session_state[ddc_match_key] = ddc_match
                                st.success(f"✅ 納品先マッチ：{ddc_match['name']}")
                                _pval = str(ddc_match.get('palette', ''))
                                _pdisp = f" / パレット：{_pval}" if _pval not in ('', 'nan', 'None') else ""
                                st.caption(f"住所：{ddc_match['address']} / TEL：{ddc_match['tel']}{_pdisp}")

                            elif ddc_result["candidates"]:
                                # --- 候補あり：ドロップダウンで選択 ---
                                st.warning(f"⚠️ 納品先「{dest_name}」の完全一致が見つかりません。候補から選択してください。")

                                candidate_options = [
                                    f"{c['name']}\uff08類似度：{c['score']:.0%}\uff09"
                                    for c in ddc_result["candidates"]
                                ]
                                candidate_options.append("\u270b 該当なし（手動入力）")

                                selected_idx = st.selectbox(
                                    "納品先候補",
                                    range(len(candidate_options)),
                                    format_func=lambda i: candidate_options[i],
                                    key=f"ddc_cand_{ocr_key}",
                                )

                                if selected_idx < len(ddc_result["candidates"]):
                                    chosen = ddc_result["candidates"][selected_idx]
                                    ddc_match = chosen["row_data"]
                                    ddc_match["matched"] = True
                                    st.session_state[ddc_match_key] = ddc_match
                                    st.info(f"📍 選択中：{chosen['name']}")
                                    _pval2 = str(ddc_match.get('palette', ''))
                                    _pdisp2 = f" / パレット：{_pval2}" if _pval2 not in ('', 'nan', 'None') else ""
                                    st.caption(f"住所：{ddc_match['address']} / TEL：{ddc_match['tel']}{_pdisp2}")
                                else:
                                    # 「該当なし」→ 手動入力 + 再検索 + 新規登録
                                    st.markdown("**納品先情報を手動入力：**")
                                    m_col1, m_col2 = st.columns(2)
                                    manual_name = m_col1.text_input("納品先名", value=dest_name, key=f"mn_{ocr_key}")
                                    manual_postal = m_col2.text_input("郵便番号", key=f"mp_{ocr_key}")
                                    m_col3, m_col4 = st.columns(2)
                                    manual_address = m_col3.text_input("住所", key=f"ma_{ocr_key}")
                                    manual_tel = m_col4.text_input("電話番号", key=f"mt_{ocr_key}")
                                    m_col5, m_col6 = st.columns(2)
                                    manual_fax = m_col5.text_input("FAX番号", key=f"mf_{ocr_key}")
                                    manual_time = m_col6.text_input("入荷時間", key=f"mti_{ocr_key}")

                                    # 再検索ボタン：手動入力した名前でDDCマスタを再検索
                                    btn_col1, btn_col2 = st.columns(2)
                                    if btn_col1.button("🔍 この名前で再検索", key=f"re_search_{ocr_key}"):
                                        re_result = match_ddc_candidates(manual_name, ddc)
                                        if re_result["exact_match"]:
                                            st.session_state[ddc_match_key] = re_result["matched_row"]
                                            st.session_state[ddc_match_key]["matched"] = True
                                            st.success(f"✅ 再検索マッチ：{re_result['matched_row']['name']}")
                                            st.rerun()
                                        elif re_result["candidates"]:
                                            st.info("💡 候補が見つかりました。上のドロップダウンで候補を選択してください。")
                                            for rc in re_result["candidates"][:3]:
                                                st.caption(f"  ・ {rc['name']}\uff08類似度：{rc['score']:.0%}\uff09")
                                        else:
                                            st.warning("検索結果：該当なし。新規登録してください。")

                                    # 新規登録ボタン：DDCマスタにCSV追加
                                    if btn_col2.button("💾 納品先マスタに新規登録", key=f"save_ddc_{ocr_key}"):
                                        if manual_name.strip():
                                            new_row = {
                                                "name": manual_name.strip(),
                                                "postal": manual_postal,
                                                "address": manual_address,
                                                "tel": manual_tel,
                                                "fax": manual_fax,
                                                "time": manual_time,
                                                "berse": "",
                                                "palette": "",
                                                "jpr": "",
                                                "method": "",
                                            }
                                            result = save_new_ddc(new_row)
                                            if result["success"]:
                                                st.success(f"✅ 「{manual_name}」をマスタに登録しました")
                                                get_ddc_master.clear()
                                                st.rerun()
                                            else:
                                                st.error(f"登録失敗：{result['message']}")
                                        else:
                                            st.error("納品先名を入力してください")

                                    ddc_match = {
                                        "matched": True,
                                        "name": manual_name,
                                        "postal": manual_postal,
                                        "address": manual_address,
                                        "tel": manual_tel,
                                        "fax": manual_fax,
                                        "time": manual_time,
                                        "berse": "無",
                                        "palette": "",
                                        "jpr": "",
                                        "method": "",
                                    }
                                    st.session_state[ddc_match_key] = ddc_match

                            else:
                                # --- 候補0件：手動入力 + 再検索 + 新規登録 ---
                                st.warning(f"⚠️ 納品先「{dest_name}」に該当する候補が見つかりません。手動で入力してください。")
                                st.markdown("**納品先情報を手動入力：**")
                                m_col1, m_col2 = st.columns(2)
                                manual_name = m_col1.text_input("納品先名", value=dest_name, key=f"mn_{ocr_key}")
                                manual_postal = m_col2.text_input("郵便番号", key=f"mp_{ocr_key}")
                                m_col3, m_col4 = st.columns(2)
                                manual_address = m_col3.text_input("住所", key=f"ma_{ocr_key}")
                                manual_tel = m_col4.text_input("電話番号", key=f"mt_{ocr_key}")
                                m_col5, m_col6 = st.columns(2)
                                manual_fax = m_col5.text_input("FAX番号", key=f"mf_{ocr_key}")
                                manual_time = m_col6.text_input("入荷時間", key=f"mti_{ocr_key}")

                                btn_col1, btn_col2 = st.columns(2)
                                if btn_col1.button("🔍 この名前で再検索", key=f"re_search_{ocr_key}"):
                                    re_result = match_ddc_candidates(manual_name, ddc)
                                    if re_result["exact_match"]:
                                        st.session_state[ddc_match_key] = re_result["matched_row"]
                                        st.session_state[ddc_match_key]["matched"] = True
                                        st.success(f"✅ 再検索マッチ：{re_result['matched_row']['name']}")
                                        st.rerun()
                                    elif re_result["candidates"]:
                                        st.info("💡 候補が見つかりました：")
                                        for rc in re_result["candidates"][:3]:
                                            st.caption(f"  ・ {rc['name']}\uff08類似度：{rc['score']:.0%}\uff09")
                                        st.caption("↑ 上の納品先欄で名前を修正して再度読み取りを実行してください")
                                    else:
                                        st.warning("検索結果：該当なし。新規登録してください。")

                                if btn_col2.button("💾 納品先マスタに新規登録", key=f"save_ddc_{ocr_key}"):
                                    if manual_name.strip():
                                        new_row = {
                                            "name": manual_name.strip(),
                                            "postal": manual_postal,
                                            "address": manual_address,
                                            "tel": manual_tel,
                                            "fax": manual_fax,
                                            "time": manual_time,
                                            "berse": "",
                                            "palette": "",
                                            "jpr": "",
                                            "method": "",
                                        }
                                        result = save_new_ddc(new_row)
                                        if result["success"]:
                                            st.success(f"✅ 「{manual_name}」をマスタに登録しました")
                                            get_ddc_master.clear()
                                            st.rerun()
                                        else:
                                            st.error(f"登録失敗：{result['message']}")
                                    else:
                                        st.error("納品先名を入力してください")

                                ddc_match = {
                                    "matched": True,
                                    "name": manual_name,
                                    "postal": manual_postal,
                                    "address": manual_address,
                                    "tel": manual_tel,
                                    "fax": manual_fax,
                                    "time": manual_time,
                                    "berse": "無",
                                    "palette": "",
                                    "jpr": "",
                                    "method": "",
                                }
                                st.session_state[ddc_match_key] = ddc_match


                            # session_stateから最新のddc_matchを取得
                            ddc_match = st.session_state.get(ddc_match_key, {"matched": False, "name": dest_name})

                            # ============================================================
                            # 商品マッチング（完全一致 → 候補選択 → 手動入力）
                            # エラーが発生しても絶対に止まらない設計
                            # ============================================================
                            def _show_product_row(st_obj, md, ocr_qty, ocr_key, idx, matched_items, key_prefix="qty"):
                                """商品行の表示・数量入力・matched_items追加を行う共通関数"""
                                dest = md["output_dest"]
                                color = "🔵" if dest == "ハルナ" else "🟤"
                                col_info, col_price, col_qty_input, col_amt = st_obj.columns([4, 1, 1, 2])
                                col_info.markdown(f"{color} **{md['master_name']}**")
                                col_price.markdown(f"CS単価  \n¥{md['cs_price']:,.0f}")
                                edited_qty = col_qty_input.number_input(
                                    "CS数", min_value=0, value=ocr_qty, step=1,
                                    key=f"{key_prefix}_{ocr_key}_{idx}",
                                )
                                amount = edited_qty * md["cs_price"]
                                col_amt.markdown(f"金額 ¥{amount:,.0f}  \n→ **{dest}**")
                                if edited_qty != ocr_qty:
                                    col_qty_input.caption(f"(OCR: {ocr_qty})")
                                matched_items.append({
                                    "matched": True, "ocr_name": md.get("ocr_name", ""),
                                    "master_name": md["master_name"], "jan": md["jan"],
                                    "code": md["code"], "spec": md["spec"], "pack": md["pack"],
                                    "unit_price": md["unit_price"], "cs_price": md["cs_price"],
                                    "output_dest": md["output_dest"],
                                    "quantity": edited_qty, "amount": amount,
                                })

                            def _manual_product_select(st_obj, pm_df, pd_mod, label, ocr_qty, ocr_key, idx, matched_items, key_prefix="prod_manual"):
                                """マスタ全商品ドロップダウンを表示する共通関数"""
                                all_prods = [f"{r['商品名']}（{r['出力先']}）" for _, r in pm_df.iterrows()]
                                all_prods.append("⏩ スキップ")
                                sel = st_obj.selectbox(
                                    f"商品を選択：{label}", range(len(all_prods)),
                                    format_func=lambda i, opts=all_prods: opts[i],
                                    key=f"{key_prefix}_{ocr_key}_{idx}",
                                )
                                if sel < len(pm_df):
                                    row = pm_df.iloc[sel]
                                    md = {
                                        "ocr_name": label,
                                        "master_name": str(row["商品名"]),
                                        "jan": str(row["JANコード"]) if pd_mod.notna(row["JANコード"]) else "",
                                        "code": str(row.get("商品コード", "")),
                                        "spec": str(row["規格"]),
                                        "pack": str(row["配送荷姿"]),
                                        "unit_price": float(row["1袋単価"]),
                                        "cs_price": float(row["CS単価"]),
                                        "output_dest": str(row["出力先"]),
                                    }
                                    _show_product_row(st_obj, md, ocr_qty, ocr_key, idx, matched_items, key_prefix="qty_m")
                                else:
                                    st_obj.info(f"⏩ 「{label}」をスキップしました")

                            ocr_items = ocr_result.get("items", [])
                            matched_items = []

                            st.markdown("**商品一覧：**")
                            st.caption("💡 OCR読取の数量が正しくない場合は、CS数欄で修正してください。商品が未マッチの場合は候補から選択できます。")

                            for idx, ocr_item in enumerate(ocr_items):
                                try:
                                    ocr_name = ocr_item.get("product_name", "")
                                    jan_code = ocr_item.get("jan_code", "")
                                    # 安全な数量変換
                                    try:
                                        ocr_qty = int(float(str(ocr_item.get("quantity_cs", 0)).strip() or "0"))
                                    except (ValueError, TypeError):
                                        ocr_qty = 0

                                    prod_result = match_product_candidates(ocr_name, jan_code, pm)

                                    if prod_result["exact_match"]:
                                        # --- 完全一致 ---
                                        md = prod_result["matched_data"]
                                        md["ocr_name"] = ocr_name
                                        _show_product_row(st, md, ocr_qty, ocr_key, idx, matched_items)

                                    elif prod_result["candidates"]:
                                        # --- 候補あり ---
                                        st.warning(f"⚠️ 商品「{ocr_name}」の完全一致が見つかりません。候補から選択してください。")
                                        cand_options = [
                                            f"{c['name']}（類似度：{c['score']:.0%}）"
                                            for c in prod_result["candidates"]
                                        ]
                                        cand_options.append("✋ 該当なし（マスタから選択）")
                                        sel_idx = st.selectbox(
                                            f"商品候補：{ocr_name}", range(len(cand_options)),
                                            format_func=lambda i, opts=cand_options: opts[i],
                                            key=f"prod_cand_{ocr_key}_{idx}",
                                        )
                                        if sel_idx < len(prod_result["candidates"]):
                                            md = prod_result["candidates"][sel_idx]["row_data"]
                                            md["ocr_name"] = ocr_name
                                            _show_product_row(st, md, ocr_qty, ocr_key, idx, matched_items, key_prefix="qty_c")
                                        else:
                                            _manual_product_select(st, pm, pd, ocr_name, ocr_qty, ocr_key, idx, matched_items, key_prefix="prod_cand_m")

                                    else:
                                        # --- 候補0件 ---
                                        st.warning(f"⚠️ 商品「{ocr_name}」（JAN: {jan_code or '不明'}）の候補が見つかりません。マスタから選択してください。")
                                        _manual_product_select(st, pm, pd, ocr_name, ocr_qty, ocr_key, idx, matched_items)

                                except Exception as prod_err:
                                    # --- エラー発生時も絶対に止めない ---
                                    st.error(f"⚠️ 商品{idx+1}の処理でエラーが発生しました。マスタから手動選択してください。")
                                    st.caption(f"エラー詳細：{str(prod_err)}")
                                    _manual_product_select(st, pm, pd, str(ocr_item.get("product_name", f"商品{idx+1}")), 0, ocr_key, idx, matched_items, key_prefix="prod_err")

                            # ============================================================
                            # 商品追加（OCR読取漏れ対応）
                            # ============================================================
                            st.markdown("---")
                            add_key = f"add_count_{ocr_key}"
                            if add_key not in st.session_state:
                                st.session_state[add_key] = 0

                            if st.button("➕ 商品を追加", key=f"add_btn_{ocr_key}"):
                                st.session_state[add_key] += 1

                            for add_idx in range(st.session_state[add_key]):
                                st.markdown(f"**追加商品 {add_idx + 1}：**")
                                add_products = [f"{r['商品名']}（{r['出力先']}）" for _, r in pm.iterrows()]
                                add_sel = st.selectbox(
                                    f"商品を選択",
                                    range(len(add_products)),
                                    format_func=lambda i, opts=add_products: opts[i],
                                    key=f"add_prod_{ocr_key}_{add_idx}",
                                )
                                add_row = pm.iloc[add_sel]
                                add_dest = str(add_row["出力先"])
                                add_color = "🔵" if add_dest == "ハルナ" else "🟤"
                                add_cs_price = float(add_row["CS単価"])

                                col_i, col_p, col_q, col_a = st.columns([4, 1, 1, 2])
                                col_i.markdown(f"{add_color} **{add_row['商品名']}**")
                                col_p.markdown(f"CS単価  \n¥{add_cs_price:,.0f}")
                                add_qty = col_q.number_input(
                                    "CS数", min_value=0, value=1, step=1,
                                    key=f"add_qty_{ocr_key}_{add_idx}",
                                )
                                add_amount = add_qty * add_cs_price
                                col_a.markdown(f"金額 ¥{add_amount:,.0f}  \n→ **{add_dest}**")

                                matched_items.append({
                                    "matched": True,
                                    "ocr_name": "(手動追加)",
                                    "master_name": str(add_row["商品名"]),
                                    "jan": str(add_row["JANコード"]) if pd.notna(add_row["JANコード"]) else "",
                                    "code": str(add_row.get("商品コード", "")),
                                    "spec": str(add_row["規格"]),
                                    "pack": str(add_row["配送荷姿"]),
                                    "unit_price": float(add_row["1袋単価"]),
                                    "cs_price": add_cs_price,
                                    "output_dest": add_dest,
                                    "quantity": add_qty,
                                    "amount": add_amount,
                                })

                            # STEP 4: PDF出力
                            st.markdown("---")
                            st.subheader("STEP 4：発注書PDF出力")

                            col_staff, col_msg = st.columns(2)
                            selected_staff = col_staff.selectbox("担当者", staff_names, key=f"staff_{ocr_key}")
                            irregular_msg = col_msg.text_input("イレギュラーリクエスト（任意）", key=f"msg_{ocr_key}")

                            # 出力先別に分離
                            haruna_items = [i for i in matched_items if i.get("matched") and i.get("output_dest") == "ハルナ"]
                            sylvia_items = [i for i in matched_items if i.get("matched") and i.get("output_dest") == "シルビア"]

                            col_h, col_s = st.columns(2)

                            # ハルナ出力
                            if haruna_items and col_h.button("📘 ハルナ発注書 生成", key=f"har_{ocr_key}"):
                                total_qty = sum(i["quantity"] for i in haruna_items)
                                order_data = {
                                    "order_no": order_no,
                                    "order_date": order_date,
                                    "delivery_date": delivery_date,
                                    "delivery_dest": ddc_match.get("name", dest_name),
                                    "quantity": total_qty,
                                    "remarks": irregular_msg,
                                }
                                ddc_data = {
                                    "postal": ddc_match.get("postal", ""),
                                    "address": ddc_match.get("address", ""),
                                    "tel": ddc_match.get("tel", ""),
                                    "fax": ddc_match.get("fax", ""),
                                    "time": ddc_match.get("time", ""),
                                    "berse": ddc_match.get("berse", "無"),
                                    "palette": ddc_match.get("palette", ""),
                                    "jpr": ddc_match.get("jpr", ""),
                                    "method": ddc_match.get("method", ""),
                                }
                                pdf_buf = gen_haruna_pdf(order_data, ddc_data, selected_staff)
                                st.download_button(
                                    "⬇️ ハルナ発注書ダウンロード",
                                    pdf_buf,
                                    f"ハルナ_{order_no}_{dest_name}.pdf",
                                    "application/pdf",
                                    key=f"dl_har_{ocr_key}"
                                )
                                st.success("✅ ハルナ発注書を生成しました")

                            # シルビア出力
                            if sylvia_items and col_s.button("📙 シルビア発注書 生成", key=f"syl_{ocr_key}"):
                                items_data = []
                                for i in sylvia_items:
                                    items_data.append({
                                        "jan": i.get("jan", ""),
                                        "name": i["master_name"],
                                        "spec": i["spec"],
                                        "pack": i["pack"],
                                        "unit_price": int(i["unit_price"]),
                                        "cs_price": int(i["cs_price"]),
                                        "quantity": i["quantity"],
                                        "amount": int(i["amount"]),
                                    })
                                order_data = {
                                    "order_no": order_no,
                                    "order_date": order_date,
                                    "delivery_date": delivery_date,
                                    "delivery_dest": ddc_match.get("name", dest_name),
                                    "postal": ddc_match.get("postal", ""),
                                    "address": ddc_match.get("address", ""),
                                    "tel": ddc_match.get("tel", ""),
                                    "fax": ddc_match.get("fax", ""),
                                    "remarks": irregular_msg,
                                }
                                pdf_buf = gen_sylvia_pdf(order_data, items_data, selected_staff)
                                st.download_button(
                                    "⬇️ シルビア発注書ダウンロード",
                                    pdf_buf,
                                    f"シルビア_{order_no}_{dest_name}.pdf",
                                    "application/pdf",
                                    key=f"dl_syl_{ocr_key}"
                                )
                                st.success("✅ シルビア発注書を生成しました")


# ============================================================
# 手動入力モード
# ============================================================
with tab_manual:
    st.header("✏️ 手動入力モード")
    st.caption("FAXのPDFがない場合（受注システム経由など）はこちらから入力してください。")

    col_type, col_staff_m = st.columns(2)
    output_type = col_type.selectbox("出力先", ["シルビア", "ハルナ"], key="manual_type")
    manual_staff = col_staff_m.selectbox("担当者", staff_names, key="manual_staff")

    col1, col2, col3 = st.columns(3)
    m_order_no = col1.text_input("オーダーNO", key="m_order_no")
    m_order_date = col2.date_input("発注日", key="m_order_date")
    m_delivery_date = col3.date_input("納品日", key="m_delivery_date")

    # 納品先選択
    if output_type == "ハルナ":
        ddc_names = ddc["納品先名"].tolist()
        m_dest = st.selectbox("納品先（DDCマスタ）", ddc_names, key="m_dest_haruna")
        ddc_row = ddc[ddc["納品先名"] == m_dest].iloc[0]

        _pval3 = str(ddc_row['パレット条件'])
        _pdisp3 = f" / パレット：{_pval3}" if _pval3 not in ('', 'nan', 'None') else ""
        st.caption(f"住所：{ddc_row['住所']} / TEL：{ddc_row['電話番号']}{_pdisp3}")

        m_qty = st.number_input("発注数量（CS）", min_value=1, value=10, key="m_qty")
        m_remarks = st.text_input("備考/イレギュラーリクエスト", key="m_remarks_h")

        if st.button("📘 ハルナ発注書 生成", key="m_gen_haruna"):
            order_data = {
                "order_no": m_order_no,
                "order_date": str(m_order_date),
                "delivery_date": str(m_delivery_date),
                "delivery_dest": m_dest,
                "quantity": m_qty,
                "remarks": m_remarks,
            }
            ddc_data = {
                "postal": str(ddc_row.get("郵便番号", "")),
                "address": str(ddc_row.get("住所", "")),
                "tel": str(ddc_row.get("電話番号", "")),
                "fax": str(ddc_row.get("FAX番号", "")),
                "time": str(ddc_row.get("入荷時間", "")),
                "berse": str(ddc_row.get("バース予約", "無")),
                "palette": str(ddc_row.get("パレット条件", "")),
                "jpr": str(ddc_row.get("JPRコード", "")),
                "method": str(ddc_row.get("納品方法", "")),
            }
            pdf_buf = gen_haruna_pdf(order_data, ddc_data, manual_staff)
            st.download_button(
                "⬇️ ダウンロード",
                pdf_buf,
                f"ハルナ_{m_order_no}_{m_dest}.pdf",
                "application/pdf",
                key="m_dl_haruna"
            )
            st.success("✅ ハルナ発注書を生成しました")

    else:  # シルビア
        m_dest_name = st.text_input("納品先名", key="m_dest_syl")
        col_a, col_b = st.columns(2)
        m_postal = col_a.text_input("郵便番号", key="m_postal")
        m_address = col_b.text_input("住所", key="m_address")
        col_c, col_d = st.columns(2)
        m_tel = col_c.text_input("電話番号", key="m_tel")
        m_fax = col_d.text_input("FAX番号", key="m_fax")

        # 商品選択
        st.markdown("**商品選択：**")
        sylvia_products = pm[pm["出力先"] == "シルビア"]

        manual_items = []
        for _, prod in sylvia_products.iterrows():
            col_name, col_qty = st.columns([3, 1])
            col_name.write(f"{prod['商品名']}（CS単価：¥{prod['CS単価']:,.0f}）")
            qty = col_qty.number_input("CS数", min_value=0, value=0, key=f"m_qty_{prod['商品名']}")
            if qty > 0:
                manual_items.append({
                    "jan": prod["JANコード"],
                    "name": prod["商品名"],
                    "spec": prod["規格"],
                    "pack": prod["配送荷姿"],
                    "unit_price": int(prod["1袋単価"]),
                    "cs_price": int(prod["CS単価"]),
                    "quantity": qty,
                    "amount": int(qty * prod["CS単価"]),
                })

        m_remarks_s = st.text_input("備考/イレギュラーリクエスト", key="m_remarks_s")

        if manual_items and st.button("📙 シルビア発注書 生成", key="m_gen_sylvia"):
            order_data = {
                "order_no": m_order_no,
                "order_date": str(m_order_date),
                "delivery_date": str(m_delivery_date),
                "delivery_dest": m_dest_name,
                "postal": m_postal,
                "address": m_address,
                "tel": m_tel,
                "fax": m_fax,
                "remarks": m_remarks_s,
            }
            pdf_buf = gen_sylvia_pdf(order_data, manual_items, manual_staff)
            st.download_button(
                "⬇️ ダウンロード",
                pdf_buf,
                f"シルビア_{m_order_no}_{m_dest_name}.pdf",
                "application/pdf",
                key="m_dl_sylvia"
            )
            st.success("✅ シルビア発注書を生成しました")

# ========== Sidebar ==========
with st.sidebar:
    st.header("ℹ️ システム情報")
    st.caption(f"商品マスタ：{len(pm)}件")
    st.caption(f"DDCマスタ：{len(ddc)}件")
    st.caption(f"担当者：{', '.join(staff_names)}")
    st.markdown("---")
    st.caption("Version 1.8 - OCR精度改善・発注日システム日付")
    st.caption("株式会社TWO 事業管理部")
