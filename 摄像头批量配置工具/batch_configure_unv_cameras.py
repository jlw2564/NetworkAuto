import argparse
import asyncio
import os
import sys
from datetime import datetime
from pathlib import Path
import re

import pandas as pd

try:
    from playwright.async_api import TimeoutError as PlaywrightTimeoutError
    from playwright.async_api import async_playwright
except ImportError:  # pragma: no cover
    print("未安装 playwright，请先执行: pip install playwright")
    print("安装浏览器内核: playwright install chromium")
    sys.exit(1)


DEFAULT_PASSWORD_CANDIDATES = [
    "admin!123",
    "admin@123",
    "admin！123",
    "admin＠123",
]

THAILAND_REGION_VALUE = "184"
PAL_VALUE = "0"
AUTOMATIC_VALUE = "0"
FRAME_RATE_25 = "25"
SUB_STREAM_2MP_VALUE = "1920*1080"
_BROWSER_DIR_CANDIDATES = [
    Path(__file__).resolve().parent / "ms-playwright",
    Path(__file__).resolve().parent.parent / "ms-playwright",
]
BROWSERS_DIR = str(next((p for p in _BROWSER_DIR_CANDIDATES if p.is_dir()), _BROWSER_DIR_CANDIDATES[1]))
SUCCESS_MESSAGE = "Parameter(s) set successfully."
VIDEO_SAVE_API = "/LAPI/V1.0/Channel/0/Media/Video/Streams/DetailInfos"

os.environ.setdefault("PLAYWRIGHT_BROWSERS_PATH", BROWSERS_DIR)


def clean_text(value):
    if pd.isna(value):
        return ""
    text = str(value).strip()
    return "" if text.lower() == "nan" else text


def split_candidates(text):
    normalized = (
        clean_text(text)
        .replace("；", ";")
        .replace("，", ",")
        .replace("、", ",")
        .replace("|", ",")
        .replace("/", ",")
    )
    if not normalized:
        return []
    parts = re.split(r",|;|\s+|或", normalized)
    return [part.strip() for part in parts if part.strip()]


def build_password_candidates(row, cli_candidates):
    candidates = []

    for column in ["密码", "登录密码", "候选密码", "密码1", "密码2", "备用密码"]:
        value = clean_text(row.get(column, ""))
        if value:
            candidates.extend(split_candidates(value))

    candidates.extend(cli_candidates)

    deduplicated = []
    seen = set()
    for candidate in candidates:
        if candidate not in seen:
            deduplicated.append(candidate)
            seen.add(candidate)
    return deduplicated


def bool_to_text(value):
    return "是" if value else "否"


class UNVCameraBatchConfigurator:
    def __init__(self, input_file, output_file, default_username, password_candidates, headless, concurrency):
        self.input_file = Path(input_file)
        self.output_file = Path(output_file)
        self.default_username = default_username
        self.password_candidates = password_candidates
        self.headless = headless
        self.concurrency = max(1, int(concurrency))

    async def run(self):
        df = pd.read_excel(self.input_file)
        if "IP" not in df.columns:
            raise ValueError(f"Excel 中未找到 'IP' 列，当前列名: {df.columns.tolist()}")

        self._ensure_result_columns(df)

        save_lock = asyncio.Lock()
        semaphore = asyncio.Semaphore(self.concurrency)

        async with async_playwright() as p:
            browser = await p.chromium.launch(headless=self.headless)
            try:
                tasks = [
                    asyncio.create_task(self._process_row(browser, df, index, row, semaphore, save_lock))
                    for index, row in df.iterrows()
                ]
                await asyncio.gather(*tasks)
            finally:
                await browser.close()

        print(f"处理完成，结果已保存到: {self.output_file}")

    async def _process_row(self, browser, df, index, row, semaphore, save_lock):
        ip = clean_text(row.get("IP", ""))
        if not ip:
            async with save_lock:
                df.at[index, "处理状态"] = "跳过"
                df.at[index, "处理结果"] = "IP 为空"
                self._save_progress(df)
            return

        username = clean_text(row.get("用户名", "")) or self.default_username
        password_candidates = build_password_candidates(row, self.password_candidates)
        if not password_candidates:
            async with save_lock:
                df.at[index, "处理状态"] = "失败"
                df.at[index, "处理结果"] = "没有可用的密码候选"
                df.at[index, "处理时间"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                self._save_progress(df)
            return

        async with semaphore:
            context = await browser.new_context(ignore_https_errors=True)
            page = await context.new_page()
            page.set_default_timeout(30000)
            dialog_state = {"message": "", "locked": False}

            async def handle_dialog(dialog):
                message = clean_text(dialog.message)
                dialog_state["message"] = message
                lowered = message.lower()
                if "locked" in lowered or "illegal login attempts" in lowered or "upper limit" in lowered:
                    dialog_state["locked"] = True
                await dialog.accept()

            page.on("dialog", handle_dialog)
            debug_trace = {"stage": "start", "selector": ""}

            try:
                result = await self._process_device(page, ip, username, password_candidates, debug_trace, dialog_state)
                async with save_lock:
                    self._write_result(df, index, username, result)
                    self._save_progress(df)
            except Exception as exc:
                async with save_lock:
                    df.at[index, "登录用户名"] = username
                    df.at[index, "处理状态"] = "失败"
                    df.at[index, "处理结果"] = str(exc)
                    df.at[index, "调试阶段"] = debug_trace.get("stage", "")
                    df.at[index, "调试选择器"] = debug_trace.get("selector", "")
                    df.at[index, "处理时间"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                    self._save_progress(df)
            finally:
                await page.close()
                await context.close()

    def _ensure_result_columns(self, df):
        extra_columns = [
            "登录用户名",
            "登录密码",
            "处理状态",
            "处理结果",
            "调试阶段",
            "调试选择器",
            "国家码是否处理",
            "国家码结果",
            "视频制式",
            "曝光模式",
            "主码流帧率",
            "辅码流分辨率",
            "辅码流帧率",
            "处理时间",
        ]
        for column in extra_columns:
            if column not in df.columns:
                df[column] = ""

    def _write_result(self, df, index, username, result):
        df.at[index, "登录用户名"] = username
        df.at[index, "登录密码"] = result["used_password"]
        df.at[index, "处理状态"] = "成功"
        df.at[index, "处理结果"] = result["message"]
        df.at[index, "调试阶段"] = result.get("debug_stage", "")
        df.at[index, "调试选择器"] = result.get("debug_selector", "")
        df.at[index, "国家码是否处理"] = bool_to_text(result["country_touched"])
        df.at[index, "国家码结果"] = result["country_result"]
        df.at[index, "视频制式"] = result["video_standard"]
        df.at[index, "曝光模式"] = result["exposure_mode"]
        df.at[index, "主码流帧率"] = result["main_fps"]
        df.at[index, "辅码流分辨率"] = result["sub_resolution"]
        df.at[index, "辅码流帧率"] = result["sub_fps"]
        df.at[index, "处理时间"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    def _save_progress(self, df):
        df.to_excel(self.output_file, index=False)

    def _frame(self, page):
        return page.frame_locator("#banner")

    def _frame_locator(self, page, selector):
        return self._frame(page).locator(selector)

    async def _content_frame(self, page):
        frame_element = await page.locator("#banner").element_handle()
        if frame_element is None:
            raise RuntimeError("未找到 iframe#banner")
        frame = await frame_element.content_frame()
        if frame is None:
            raise RuntimeError("无法获取 iframe#banner 的内容上下文")
        return frame

    async def _wait_visible(self, page, selector, timeout=5000, trace=None, stage=None):
        if trace is not None:
            if stage:
                trace["stage"] = stage
            trace["selector"] = selector
        try:
            await self._frame_locator(page, selector).wait_for(state="visible", timeout=timeout)
            return True
        except PlaywrightTimeoutError:
            return False

    async def _click(self, page, selector, timeout=10000, trace=None, stage=None):
        if trace is not None:
            if stage:
                trace["stage"] = stage
            trace["selector"] = selector
        locator = self._frame_locator(page, selector)
        await locator.wait_for(state="visible", timeout=timeout)
        await locator.scroll_into_view_if_needed()
        await locator.click(force=True)

    async def _option_values(self, page, selector, trace=None, stage=None):
        if trace is not None:
            if stage:
                trace["stage"] = stage
            trace["selector"] = selector
        locator = self._frame_locator(page, selector)
        await locator.wait_for(state="visible", timeout=10000)
        return await locator.evaluate(
            """el => Array.from(el.options).map(opt => ({
                value: opt.value,
                text: opt.textContent.trim()
            }))"""
        )

    async def _set_select_value(self, page, selector, target_value, field_name, trace=None, stage=None):
        locator = self._frame_locator(page, selector)
        if trace is not None:
            if stage:
                trace["stage"] = stage
            trace["selector"] = selector
        await locator.wait_for(state="visible", timeout=10000)
        await locator.scroll_into_view_if_needed()

        options = await self._option_values(page, selector, trace=trace, stage=stage)
        option_values = [item["value"] for item in options]
        if target_value not in option_values:
            available = ", ".join(f'{item["value"]}:{item["text"]}' for item in options)
            raise RuntimeError(f"{field_name} 不支持目标值 {target_value}，当前可选项: {available}")

        before_value = await locator.input_value()
        await locator.hover()
        await locator.click(force=True)
        await locator.select_option(value=target_value)
        await locator.blur()

        for _ in range(10):
            current_value = await locator.input_value()
            if current_value == target_value:
                break
            await locator.select_option(value=target_value)
            await locator.blur()
            await page.wait_for_timeout(300)

        current_value = await locator.input_value()
        if current_value != target_value:
            raise RuntimeError(
                f"{field_name} 设置失败，目标值 {target_value}，当前值 {current_value}，原值 {before_value}"
            )

    async def _wait_success_message(self, page, timeout=15000, trace=None, stage=None):
        if trace is not None:
            if stage:
                trace["stage"] = stage
            trace["selector"] = "#msg"

        end_time = asyncio.get_running_loop().time() + timeout / 1000
        last_text = ""

        while asyncio.get_running_loop().time() < end_time:
            msg = self._frame_locator(page, "#msg")
            text = clean_text(await msg.text_content() or "")
            if text:
                last_text = text
            if SUCCESS_MESSAGE in text:
                return
            await page.wait_for_timeout(300)

        raise RuntimeError(f"保存后提示异常: {last_text or '未读取到成功提示文本'}")

    async def _click_and_wait_video_save(self, page, trace=None):
        if trace is not None:
            trace["stage"] = "video_save"
            trace["selector"] = "span.saveBtn"
        try:
            async with page.expect_response(
                lambda resp: resp.request.method == "PUT" and VIDEO_SAVE_API in resp.url,
                timeout=15000,
            ) as response_info:
                await self._click(page, "span.saveBtn", trace=trace, stage="video_save_click")
            response = await response_info.value
            if not response.ok:
                raise RuntimeError(f"视频参数保存失败，HTTP {response.status}")
            await page.wait_for_timeout(1200)
            await self._wait_success_message(page, trace=trace, stage="video_save_message")
            return
        except PlaywrightTimeoutError:
            if trace is not None:
                trace["stage"] = "video_save_no_response"
                trace["selector"] = VIDEO_SAVE_API
            await page.wait_for_timeout(1200)
            return

    async def _save_image_settings(self, page, trace=None):
        if trace is not None:
            trace["stage"] = "image_save"
            trace["selector"] = "#submitBut"

        if await self._wait_visible(page, "#submitBut", timeout=1500, trace=trace, stage="image_submit_visible"):
            await self._click(page, "#submitBut", trace=trace, stage="image_submit_click")
        else:
            if trace is not None:
                trace["stage"] = "image_submit_eval"
                trace["selector"] = "Live.submitSetupArea()"
            frame = await self._content_frame(page)
            await frame.evaluate("() => { if (!window.Live || !window.Live.submitSetupArea) throw new Error('Live.submitSetupArea 不存在'); window.Live.submitSetupArea(); }")

        await page.wait_for_timeout(1200)
        await self._wait_success_message(page, trace=trace, stage="image_save_message")

    async def _goto_device(self, page, ip, trace=None):
        if trace is not None:
            trace["stage"] = "goto_device"
            trace["selector"] = "#userName"
        await page.goto(f"http://{ip}", wait_until="domcontentloaded")
        await page.wait_for_timeout(1000)
        if not await self._wait_visible(page, "#userName", timeout=15000, trace=trace, stage="login_page"):
            raise RuntimeError("登录页加载失败，未找到用户名输入框")

    async def _login(self, page, ip, username, password_candidates, trace=None, dialog_state=None):
        last_error = "登录失败"

        for password in password_candidates:
            if dialog_state is not None:
                dialog_state["message"] = ""
                dialog_state["locked"] = False
            await self._goto_device(page, ip, trace=trace)
            frame = self._frame(page)
            if trace is not None:
                trace["stage"] = "login_submit"
                trace["selector"] = "#login"
            await frame.locator("#userName").fill(username)
            await frame.locator("#password").fill("")
            await frame.locator("#password").fill(password)
            await frame.locator("#login").click()
            await page.wait_for_timeout(2500)

            if await self._wait_post_login(page, trace=trace):
                return password

            if dialog_state is not None and dialog_state.get("locked"):
                raise RuntimeError(f"账号已锁定: {dialog_state.get('message', '')}")

            if dialog_state is not None and dialog_state.get("message"):
                last_error = f"密码尝试失败: {password} | {dialog_state['message']}"
            else:
                last_error = f"密码尝试失败: {password}"

        raise RuntimeError(last_error)

    async def _wait_post_login(self, page, timeout_ms=20000, trace=None):
        elapsed = 0
        while elapsed < timeout_ms:
            if await self._wait_visible(page, "#region", timeout=1000, trace=trace, stage="post_login_region"):
                return True
            if await self._wait_visible(page, "#nav_config", timeout=1000, trace=trace, stage="post_login_setup"):
                return True
            if await self._wait_visible(page, "#imgCfgLink", timeout=1000, trace=trace, stage="post_login_image_link"):
                return True
            if await self._wait_visible(page, "#videoLink", timeout=1000, trace=trace, stage="post_login_video_link"):
                return True
            await page.wait_for_timeout(500)
            elapsed += 500
        return False

    async def _handle_initialization_if_needed(self, page, trace=None):
        if not await self._wait_visible(page, "#region", timeout=2500, trace=trace, stage="init_region_check"):
            return False, "未触发国家码初始化"

        await self._set_select_value(page, "#region", THAILAND_REGION_VALUE, "国家码", trace=trace, stage="init_region")
        if await self._wait_visible(page, "#PNMode", timeout=1500, trace=trace, stage="init_pnmode_check"):
            await self._set_select_value(page, "#PNMode", PAL_VALUE, "初始化视频制式", trace=trace, stage="init_pnmode")
        await self._click(page, "#submit", trace=trace, stage="init_submit")
        await page.wait_for_timeout(3000)
        return True, "已选择泰国并完成初始化"

    async def _enter_setup(self, page, trace=None):
        if await self._wait_visible(page, "#nav_config", timeout=6000, trace=trace, stage="setup_nav_check"):
            await self._click(page, "#nav_config", trace=trace, stage="setup_nav_click")
            await page.wait_for_timeout(1000)

    async def _ensure_menu_expanded(self, page, child_selector, menu_selector, trace=None, stage_prefix="menu"):
        if await self._wait_visible(page, child_selector, timeout=1200, trace=trace, stage=f"{stage_prefix}_child_check"):
            return
        await self._click(page, menu_selector, timeout=10000, trace=trace, stage=f"{stage_prefix}_expand")
        if not await self._wait_visible(page, child_selector, timeout=10000, trace=trace, stage=f"{stage_prefix}_child_wait"):
            raise RuntimeError(f"菜单展开失败，目标子菜单不可见: {child_selector}")

    async def _ensure_panel_expanded(self, page, panel_toggle_selector, target_selector, trace=None, stage_prefix="panel"):
        if await self._wait_visible(page, target_selector, timeout=1200, trace=trace, stage=f"{stage_prefix}_target_check"):
            return
        await self._click(page, panel_toggle_selector, timeout=10000, trace=trace, stage=f"{stage_prefix}_expand")
        if not await self._wait_visible(page, target_selector, timeout=10000, trace=trace, stage=f"{stage_prefix}_target_wait"):
            raise RuntimeError(f"分组展开失败，目标控件不可见: {target_selector}")

    async def _open_image_page(self, page, trace=None):
        await self._enter_setup(page, trace=trace)
        await self._ensure_menu_expanded(
            page,
            "#imgCfgLink",
            "div.menu-title-text:has(#img)",
            trace=trace,
            stage_prefix="image_menu",
        )
        await self._click(page, "#imgCfgLink", timeout=15000, trace=trace, stage="image_link_click")
        await self._frame_locator(page, "#imgConfigTab").wait_for(state="visible", timeout=20000)
        await self._frame_locator(page, "#VideoStandardLeg").wait_for(state="visible", timeout=10000)

    async def _open_video_page(self, page, trace=None):
        await self._enter_setup(page, trace=trace)
        await self._ensure_menu_expanded(
            page,
            "#videoLink",
            "div.menu-title-text:has(#audio_video)",
            trace=trace,
            stage_prefix="video_menu",
        )
        await self._click(page, "#videoLink", timeout=15000, trace=trace, stage="video_link_click")
        await self._frame_locator(page, "#MainStreamFrameRate").wait_for(state="visible", timeout=20000)

    async def _apply_image_settings(self, page, trace=None):
        await self._open_image_page(page, trace=trace)
        await self._ensure_panel_expanded(
            page,
            "#VideoStandardLeg",
            "#PNMode",
            trace=trace,
            stage_prefix="image_video_standard",
        )
        await self._ensure_panel_expanded(
            page,
            "#ExposureLeg",
            "#ExposureMode",
            trace=trace,
            stage_prefix="image_exposure_panel",
        )
        await self._set_select_value(page, "#PNMode", PAL_VALUE, "视频制式", trace=trace, stage="image_pnmode")
        await self._set_select_value(page, "#ExposureMode", AUTOMATIC_VALUE, "曝光模式", trace=trace, stage="image_exposure")
        await self._save_image_settings(page, trace=trace)

    async def _apply_video_settings(self, page, trace=None):
        await self._open_video_page(page, trace=trace)
        await self._set_select_value(page, "#MainStreamFrameRate", FRAME_RATE_25, "主码流帧率", trace=trace, stage="video_main_fps")
        await page.wait_for_timeout(300)
        await self._set_select_value(page, "#AuxStreamPictureSize", SUB_STREAM_2MP_VALUE, "辅码流分辨率", trace=trace, stage="video_sub_resolution")
        await page.wait_for_timeout(600)
        await self._set_select_value(page, "#AuxStreamFrameRate", FRAME_RATE_25, "辅码流帧率", trace=trace, stage="video_sub_fps")
        await page.wait_for_timeout(300)
        await self._click_and_wait_video_save(page, trace=trace)

    async def _verify_values(self, page, trace=None):
        await self._open_image_page(page, trace=trace)
        await self._ensure_panel_expanded(
            page,
            "#VideoStandardLeg",
            "#PNMode",
            trace=trace,
            stage_prefix="verify_video_standard",
        )
        await self._ensure_panel_expanded(
            page,
            "#ExposureLeg",
            "#ExposureMode",
            trace=trace,
            stage_prefix="verify_exposure_panel",
        )
        await self._frame_locator(page, "#PNMode").wait_for(state="visible", timeout=10000)
        video_standard = await self._frame_locator(page, "#PNMode").input_value()
        exposure_mode = await self._frame_locator(page, "#ExposureMode").input_value()

        await self._open_video_page(page, trace=trace)
        main_fps = await self._frame_locator(page, "#MainStreamFrameRate").input_value()
        sub_resolution = await self._frame_locator(page, "#AuxStreamPictureSize").input_value()
        sub_fps = await self._frame_locator(page, "#AuxStreamFrameRate").input_value()

        return {
            "video_standard": "PAL" if video_standard == PAL_VALUE else video_standard,
            "exposure_mode": "Automatic" if exposure_mode == AUTOMATIC_VALUE else exposure_mode,
            "main_fps": main_fps,
            "sub_resolution": sub_resolution,
            "sub_fps": sub_fps,
        }

    def _is_target_config(self, values):
        return (
            values["video_standard"] == "PAL"
            and values["exposure_mode"] == "Automatic"
            and values["main_fps"] == FRAME_RATE_25
            and values["sub_resolution"] == SUB_STREAM_2MP_VALUE
            and values["sub_fps"] == FRAME_RATE_25
        )

    async def _process_device(self, page, ip, username, password_candidates, trace, dialog_state):
        used_password = await self._login(page, ip, username, password_candidates, trace=trace, dialog_state=dialog_state)
        country_touched, country_result = await self._handle_initialization_if_needed(page, trace=trace)
        current_values = await self._verify_values(page, trace=trace)
        if self._is_target_config(current_values):
            return {
                "used_password": used_password,
                "debug_stage": trace.get("stage", ""),
                "debug_selector": trace.get("selector", ""),
                "country_touched": country_touched,
                "country_result": country_result,
                "video_standard": current_values["video_standard"],
                "exposure_mode": current_values["exposure_mode"],
                "main_fps": current_values["main_fps"],
                "sub_resolution": "1920x1080(2MP)" if current_values["sub_resolution"] == SUB_STREAM_2MP_VALUE else current_values["sub_resolution"],
                "sub_fps": current_values["sub_fps"],
                "message": "参数已符合要求，无需修改",
            }
        await self._apply_image_settings(page, trace=trace)
        await self._apply_video_settings(page, trace=trace)
        final_values = await self._verify_values(page, trace=trace)

        if not self._is_target_config(final_values):
            if final_values["video_standard"] != "PAL":
                raise RuntimeError(f"视频制式校验失败，当前值: {final_values['video_standard']}")
            if final_values["exposure_mode"] != "Automatic":
                raise RuntimeError(f"曝光模式校验失败，当前值: {final_values['exposure_mode']}")
            if final_values["main_fps"] != FRAME_RATE_25:
                raise RuntimeError(f"主码流帧率校验失败，当前值: {final_values['main_fps']}")
            if final_values["sub_resolution"] != SUB_STREAM_2MP_VALUE:
                raise RuntimeError(f"辅码流分辨率校验失败，当前值: {final_values['sub_resolution']}")
            if final_values["sub_fps"] != FRAME_RATE_25:
                raise RuntimeError(f"辅码流帧率校验失败，当前值: {final_values['sub_fps']}")

        return {
            "used_password": used_password,
            "debug_stage": trace.get("stage", ""),
            "debug_selector": trace.get("selector", ""),
            "country_touched": country_touched,
            "country_result": country_result,
            "video_standard": final_values["video_standard"],
            "exposure_mode": final_values["exposure_mode"],
            "main_fps": final_values["main_fps"],
            "sub_resolution": "1920x1080(2MP)" if final_values["sub_resolution"] == SUB_STREAM_2MP_VALUE else final_values["sub_resolution"],
            "sub_fps": final_values["sub_fps"],
            "message": "参数修改并校验成功",
        }


def parse_args():
    parser = argparse.ArgumentParser(description="宇视摄像头 Excel 批量配置工具")
    parser.add_argument(
        "-i",
        "--input",
        default="camera_batch.xlsx",
        help="输入 Excel 路径，必须包含 IP 列",
    )
    parser.add_argument(
        "-o",
        "--output",
        default="camera_batch_results0722.xlsx",
        help="输出 Excel 路径",
    )
    parser.add_argument(
        "-u",
        "--username",
        default="admin",
        help="默认登录用户名，Excel 未提供用户名时使用",
    )
    parser.add_argument(
        "-p",
        "--passwords",
        default=",".join(DEFAULT_PASSWORD_CANDIDATES),
        help="默认密码候选，多个密码可用逗号分隔",
    )
    parser.add_argument(
        "--headless",
        action="store_true",
        help="启用无头模式运行浏览器",
    )
    parser.add_argument(
        "-c",
        "--concurrency",
        type=int,
        default=10,
        help="并发处理数量，默认 10",
    )
    return parser.parse_args()


async def main():
    args = parse_args()
    password_candidates = split_candidates(args.passwords)

    configurator = UNVCameraBatchConfigurator(
        input_file=args.input,
        output_file=args.output,
        default_username=args.username,
        password_candidates=password_candidates,
        headless=args.headless,
        concurrency=args.concurrency,
    )
    await configurator.run()


if __name__ == "__main__":
    try:
        asyncio.run(main())
    except KeyboardInterrupt:
        print("用户中断运行")
        sys.exit(1)
    except Exception as exc:
        print(f"运行失败: {exc}")
        if "Executable doesn't exist" in str(exc) or "browserType.launch" in str(exc):
            print("请先安装 Playwright 浏览器内核，命令如下：")
            print(f"$env:PLAYWRIGHT_BROWSERS_PATH='{BROWSERS_DIR}'; python -m playwright install chromium")
        sys.exit(1)
