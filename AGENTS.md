# Working on WB analytics

- User prefers simple implementation, minimal manual actions and targeted checks
  only for actual changes/problems. Do not claim live deployment from a Git push.
- Never commit API keys, receiver tokens, browser profiles/cookies, cost records,
  Google credentials, user spreadsheets or server `.env` files.
- Preserve manual values and formulas. The comparison tab is
  `Прогноз по конверсиям`; ordinary source updates use `Рабочая табличка`.
- Source for the receiver UI is `src/Comparison.gs` and
  `src/ComparisonSidebar.html`. Root `CostReceiver.gs` is the assembled receiver;
  its base protocol above the generated marker is edited directly.
- Forecast W/X are reserved for conversion-only ABC and relative index deviation.
  The private cost helper moves right of Z and must be found by its three headers,
  never hardcoded to X:Z. Preserve costs when migrating. Financial values must not
  influence ABC; incomplete funnels must not be converted to zero scores.
- Before publishing code updates, run `python3 tools/build_all_releases.py`.
  Commit affected source, manifests, and ALL generated immutable release files
  together. Preserve old `releases/`, `receiver-releases/`, `server-releases/`.
- `Loader.gs` and `ReceiverLoader.gs` are permanent installed entry points.
  Ordinary changes go into remote release sources, not these installed files.
  New sidebar RPCs can use `wbReceiverRpc(action, argument)` with explicit exports
  in `tools/build_comparison_receiver.py`; never expose arbitrary evaluation.
- Server updater has a fixed target list. New dependencies, permission scopes,
  changes to the installed updater or permanent loader require explicit setup;
  do not claim such changes become automatic merely by changing the repository.
- A successful source check is not proof of WB authorization, server installation
  or spreadsheet writes. Report observed status and missing verification clearly.
