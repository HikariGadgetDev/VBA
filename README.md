## Design Notes

This macro is intentionally defensive.

- Rejects multi-area selections to avoid Excel-side corruption  
- Does not delete columns by default (destructive operation)  
- Handles error cells (#N/A), Null, Empty, and numeric values safely  
- Restores Excel application state on all exit paths  

The goal is not brevity, but operational safety in real-world Excel environments.

---

## 設計メモ（日本語）

本マクロは **意図的に防御的な設計**を採用しています。

- 複数範囲選択（Ctrlによる飛び飛び選択）を拒否し、  
  Excel 側のデータ破損を防止します  
- 破壊的操作である「列削除」はデフォルトでは行いません  
- エラーセル（#N/A 等）、Null、Empty、数値が混在していても  
  処理が中断・破損しないよう安全に扱います  
- すべての終了経路で、Excel のアプリケーション状態  
  （画面更新・計算・イベント等）を確実に復元します  

本マクロの目的はコードの短さではなく、  
**実運用環境における安全性と信頼性**です。
