# Specification documents

`docs/specs/` は現在有効なcontract、invariant、compatibility rule、validation requirementだけを保持する。判断理由と却下案はADRへ、作業予定は `tasks/todo.md` へ置く。

- 仕様変更と実装・testは同じchange setで更新する。
- 過去仕様を本文へ蓄積せず、user-visibleな変更履歴はCHANGELOGへ移す。
- public API、error behavior、security boundary、release validationにはauthoritativeなtestまたは検査scriptへの参照を含める。
