# excel-sheet-navigation
Excel VBA macro for switching between sheets using keyboard shortcuts — ideal for laptops and compact keyboards without Page Up / Page Down keys.
# Excel Sheet Navigation VBA

## 🔍 Overview

This lightweight and portable VBA macro lets you navigate between sheets in any Excel workbook using keyboard shortcuts — a perfect solution for users on **laptops**, **compact keyboards**, or **devices that lack Page Up / Page Down keys**.

## Why Use This?

Many laptops and smaller keyboards don't include dedicated navigation keys like `Page Up` or `Page Down`, making Excel tab switching cumbersome. This macro lets you:

- Use `Ctrl + Shift + N` to go to the **next sheet**
- Use `Ctrl + Shift + B` to go to the **previous sheet**

No need to touch your mouse or click on tiny arrows again!

## 📦 How to Import the VBA Module

1. Open your Excel workbook.
2. Press `Alt + F11` to open the **VBA Editor**.
3. In the editor, go to `File > Import File...`.
4. Select the downloaded `SheetNavigation.bas` file.
5. The macro will be added under `Modules` as `SheetNavigation`.

---

## 🎯 How to Assign Shortcut Keys

1. Go back to Excel.
2. Press `Alt + F8` to open the **Macro window**.
3. Select `GoToNextSheet` → click **Options** → set shortcut to `Ctrl + Shift + N`.
4. Select `GoToPreviousSheet` → click **Options** → set shortcut to `Ctrl + Shift + B`.
5. Click OK. Your shortcuts are now active.

> ⚠️ Note: Excel does not support assigning shortcuts using arrow keys (↑/↓), so `N` and `B` are used instead.

---

## 📝 License

MIT License — see LICENSE file for details. Feel free to use, modify, and distribute.
