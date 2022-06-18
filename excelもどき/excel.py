
import dearpygui.dearpygui as dpg
import re

# ===================================================
#  imguiで表示
# ===================================================
dpg.create_context()
dpg.create_viewport(title='Excel @ python', width=720, height=490, clear_color=(69, 132, 184))

# ==== フォントの設定 ====================================
with dpg.font_registry():
    with dpg.font('PlemolJP35ConsoleNF-Text.ttf', 15, default_font=True) as DEF_FONT:
        dpg.add_font_range_hint(dpg.mvFontRangeHint_Chinese_Full)
        dpg.add_font_range(0x2000, 0x2e60)
        dpg.add_font_range(0xe000, 0xf8ff)
dpg.bind_font(DEF_FONT)


# ==== 全体のカラーテーマ =================================
with dpg.theme() as all_theme:
    with dpg.theme_component(dpg.mvAll):
        # ウィンドウの背景色
        dpg.add_theme_color(dpg.mvThemeCol_WindowBg, (255, 255, 255))
        # ウィンドウの背景色
        dpg.add_theme_color(dpg.mvThemeCol_Border, (150, 150, 150))
        # タイトルバーの背景色
        dpg.add_theme_color(dpg.mvThemeCol_TitleBg, (240, 240, 240))
        # タイトルバーの選択時の背景色
        dpg.add_theme_color(dpg.mvThemeCol_TitleBgActive, (240, 240, 240))
        # タイトルバーの選択時の背景色
        dpg.add_theme_color(dpg.mvThemeCol_TitleBgCollapsed, (240, 240, 240))
        # テキストの色
        dpg.add_theme_color(dpg.mvThemeCol_Text, (50, 50, 50))
        # メニューバーの背景色
        dpg.add_theme_color(dpg.mvThemeCol_MenuBarBg, (255, 255, 255))
        # メニューエリアの背景色
        dpg.add_theme_color(dpg.mvThemeCol_PopupBg, (255, 255, 255))
        # メニューバーの選択状態の色
        dpg.add_theme_color(dpg.mvThemeCol_Header, (143, 207, 245, 179))
        # メニューのホバーの色
        dpg.add_theme_color(dpg.mvThemeCol_HeaderHovered, (143, 207, 245, 102))
        # メニューバーのアクティブ状態の色
        dpg.add_theme_color(dpg.mvThemeCol_HeaderActive, (143, 207, 245, 200))
        # メニューバーの左右、メニューアイテムの左右上下
        dpg.add_theme_style(dpg.mvStyleVar_WindowPadding, 8, 5)
        # メニューバーの上下のみ
        dpg.add_theme_style(dpg.mvStyleVar_FramePadding, 0, 2)
        # メニューのボーダーを消す
        dpg.add_theme_style(dpg.mvStyleVar_PopupBorderSize, 1)
        # チェックの背景色
        # dpg.add_theme_color(dpg.mvThemeCol_FrameBg, (220, 220, 220))
        dpg.add_theme_color(dpg.mvThemeCol_FrameBg, (232, 234, 237))
        # チェックのホバー時の背景色
        dpg.add_theme_color(dpg.mvThemeCol_FrameBgHovered, (143, 207, 245, 102))
        # チェックのaアクティブ時の背景色
        dpg.add_theme_color(dpg.mvThemeCol_FrameBgActive, (143, 207, 245, 179))
        # フレームの角
        dpg.add_theme_style(dpg.mvStyleVar_FrameRounding, 3)
        # グリップの背景色
        dpg.add_theme_color(dpg.mvThemeCol_ResizeGrip, (255, 255, 255, 0))
        # グリップの背景色
        dpg.add_theme_color(dpg.mvThemeCol_ResizeGripHovered, (255, 255, 255, 0))
        # グリップの背景色
        dpg.add_theme_color(dpg.mvThemeCol_ResizeGripActive, (255, 255, 255, 0))
        # スクロールバーの背景色
        dpg.add_theme_color(dpg.mvThemeCol_ScrollbarBg, (255, 255, 255, 0))
        # スクロールバーの背景色
        dpg.add_theme_color(dpg.mvThemeCol_ScrollbarGrab, (150, 150, 150))
        # チャイルドウィンドウの背景色
        dpg.add_theme_color(dpg.mvThemeCol_ChildBg, (255, 255, 255))
        # チャイルドウィンドウの左右上下
        dpg.add_theme_style(dpg.mvStyleVar_WindowPadding, 4, 4)
        # タブの背景色
        # dpg.add_theme_color(dpg.mvThemeCol_Tab, (230, 240, 255))
        dpg.add_theme_color(dpg.mvThemeCol_Tab, (232, 234, 237))
        # タブのホバー時の背景色
        dpg.add_theme_color(dpg.mvThemeCol_TabHovered, (143, 207, 245, 102))
        # dpg.add_theme_color(dpg.mvThemeCol_TabHovered, (180, 180, 180))
        # タブのアクティブ時の背景色
        # dpg.add_theme_color(dpg.mvThemeCol_TabActive, (143, 207, 245))
        dpg.add_theme_color(dpg.mvThemeCol_TabActive, (143, 207, 245, 179))
        # ボタンの背景色
        dpg.add_theme_color(dpg.mvThemeCol_Button, (143, 207, 245, 179))
        # ボタンの背景色
        dpg.add_theme_color(dpg.mvThemeCol_ButtonHovered, (143, 207, 245, 102))
        # ボタンの背景色
        dpg.add_theme_color(dpg.mvThemeCol_ButtonActive, (143, 207, 245, 200))

        # テーブルのヘッダ
        dpg.add_theme_color(dpg.mvThemeCol_TableHeaderBg,  (220, 220, 220))
dpg.bind_theme(all_theme)

# ==== 背景とメニューのテーマ ================================
with dpg.theme() as background_window_theme:
    with dpg.theme_component(dpg.mvWindowAppItem):
        # ウィンドウの背景色
        dpg.add_theme_color(dpg.mvThemeCol_WindowBg, (240, 240, 240))
        # ウィンドウのボーダーを消す
        dpg.add_theme_style(dpg.mvStyleVar_WindowBorderSize, 0)


# ==== メインウィンドウ ======================================
with dpg.window(label='bgwindow') as background_window:
    # ==== メニュー ============================================
    with dpg.menu_bar():
        with dpg.menu(label='ファイル'):
            dpg.add_menu_item(label='\uf07c  開けない', shortcut='   Ctrl+O')
    # ==== メニュー ============================================
    with dpg.group(horizontal=True):
        with dpg.child_window(width=-1,border=True):
            display_text = None
            inner_formula = None
            with dpg.group(horizontal=True):
                dpg.add_text(' 値:')
                display_text = dpg.add_input_text(width=200,readonly=True)
                dpg.add_text(' 式:')
                inner_formula = dpg.add_input_text(width=-1)
            dpg.add_separator()
            dpg.push_container_stack(dpg.add_table(resizable=True, borders_innerH=True, borders_outerH=True, borders_innerV=True, borders_outerV=True))
            columns = 12
            rows = 16
            cells = []
            before_select = None

            def update(taishou_cell):
                userdata =  dpg.get_item_user_data(taishou_cell)
                formula = userdata['formula']
                display_data = ''
                if len(formula) != 0:
                    if formula[0] == '=':
                        try:
                            display_data = eval(re.sub(r'([a-zA-Z]\d+)',r"c('\1')",formula[1:]))
                        except:
                            display_data = '[ERR]'
                    else:
                        display_data = formula
                else:
                    display_data = formula
                dpg.set_item_label(taishou_cell, display_data)
                dpg.set_item_user_data(taishou_cell, {'cells':cells,'formula':formula, 'display':display_data})

            def c(cell_str):
                global cells
                index = ord(cell_str[0].upper())-65
                index += (int(cell_str[1:])-1)*(columns-1)
                ss = dpg.get_item_user_data(cells[index])['display']
                
                if type(ss) is str:
                    if re.match(r'^\d+$',ss):
                        return int(ss)
                    elif re.match(r'^\d+\.\d*$',ss):
                        return float(ss)
                return ss

            def _cell_select(sender, app_data, user_data):
                global before_select
                if before_select != None:
                    formula = dpg.get_value(inner_formula)
                    display_data = ''
                    if len(formula) != 0:
                        if formula[0] == '=':
                            try:
                                display_data = eval(re.sub(r'([a-zA-Z]\d+)',r"c('\1')",formula[1:]))
                            except:
                                display_data = '[ERR]'
                        else:
                            display_data = formula
                    else:
                        display_data = formula
                    dpg.set_item_label(before_select, display_data )
                    dpg.set_item_user_data(before_select, {'cells':cells,'formula':formula, 'display':display_data})
                dpg.set_value(display_text, user_data['display'])
                dpg.set_value(inner_formula, user_data['formula'])
                for item in user_data['cells']:
                    if item != sender:
                        dpg.set_value(item, False)
                        update(item)
                    else:
                        dpg.set_value(item, True)
                dpg.focus_item(inner_formula)
                before_select = sender
            
            dpg.add_table_column(label='#',width_fixed=True)
            for i in range(1,columns):
                dpg.add_table_column(label=chr(64+i),width_stretch=True)
            for i in range(rows):
                with dpg.table_row():
                    dpg.add_text(str(i+1).rjust(3,' '))
                    for j in range(1, columns):
                        # dpg.add_text('hoge')
                        cells.append(dpg.add_selectable(label="", callback=_cell_select, user_data={'cells':cells,'formula':'', 'display':''}))
            dpg.pop_container_stack()



dpg.set_primary_window(background_window, True)
dpg.bind_item_theme(background_window, background_window_theme)



dpg.setup_dearpygui()
dpg.show_viewport()
dpg.start_dearpygui()
dpg.destroy_context()