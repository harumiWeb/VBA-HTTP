# Architecture Decision Records

ADRは、後続の実装者が判断理由とtrade-offを必要とする長期的な設計判断を記録する。現在の振る舞いや手順は `docs/specs/` に置く。

## Naming and status

- File name: `ADR-NNNN-short-kebab-title.md`
- Number: four-digit, monotonically increasing, and never reused
- Status: `proposed`, `accepted`, `superseded`, or `deprecated`
- Language: repository documentationに合わせる。technical identifierとcommandは原表記を維持する。

## Required sections

1. Status
2. Background
3. Decision
4. Consequences
5. Rationale（tests、code、related specsへの証跡）
6. Supersedes
7. Superseded by

## Lifecycle

- 実装前にADRを `proposed` として作成し、合意済み判断は `accepted` にする。
- 受理済みADRの判断を静かに書き換えない。変更する場合は新しいADRを作成する。
- 新しいADRは旧ADRの `Superseded by` を更新し、旧ADRを `superseded` にする。新ADRの `Supersedes` から旧ADRを参照する。
- 実装、spec、testがADRと一致しない場合は、実装を直すか新しい判断をADR化するまで完了扱いにしない。
