# Code signing policy · 코드 서명 정책

> **Status / 상태:** Office Kit plans to apply for free code signing from SignPath Foundation. Until an application is approved,
> releases are **not signed** and Windows SmartScreen may warn on first run.
> SignPath Foundation에 무료 코드 서명을 신청할 예정입니다. 승인 전까지 Release는 **서명되지 않으며**, 처음 실행할 때 SmartScreen 경고가 뜰 수 있습니다.

Free code signing provided by [SignPath.io](https://about.signpath.io/), certificate by [SignPath Foundation](https://signpath.org/).
(Applies once an application is approved. / 신청이 승인되면 적용됩니다.)

## What is signed / 서명하는 것

- The Windows installer `OfficeKit-Setup-<version>.exe` and the `Office Kit.exe` it installs, built by GitHub Actions
  ([`.github/workflows/release.yml`](../.github/workflows/release.yml)) from the tagged source in this repository.
- Upstream open-source binaries included in the installer (the Electron runtime libraries, the rhwp converter
  `rhwp.exe`, the office editor files) are shipped **as published by their projects** and are not signed with this
  project's certificate.
- Every signing request is approved manually by an approver (below). Builds come only from GitHub-hosted runners.

## Team roles / 역할

Office Kit is maintained by one person. / 한 사람이 관리합니다.

| Role | Members |
|---|---|
| Committers and reviewers | [@toxh13](https://github.com/toxh13) (repository owner) |
| Approvers | [@toxh13](https://github.com/toxh13) (repository owner) |

Changes proposed by anyone else (pull requests) are reviewed by a committer before they are merged.
Every team member must use multi-factor authentication for GitHub and SignPath.

## Privacy policy / 개인정보

This program will not transfer any information to other networked systems unless specifically requested by the user
or the person installing or operating it.

이 프로그램은 사용자(또는 설치·운영하는 사람)가 직접 요청하지 않는 한, 어떤 정보도 다른 네트워크 시스템으로 보내지 않습니다.

What happens only when the user asks for it / 사용자가 요청할 때만 일어나는 일:

| When / 언제 | What is sent, to whom / 무엇을 어디로 |
|---|---|
| The user sends a message in the AI chat | The document's content and the message, through the Claude Code or Codex CLI the user installed and signed in to, to **Anthropic** ([privacy policy](https://www.anthropic.com/legal/privacy)) or **OpenAI** ([privacy policy](https://openai.com/policies/privacy-policy/)) under the user's own account |
| The user clicks "connect Hangul automation" | A download of Hancom's security module from Hancom's official GitHub repository |
| The user opens a link in a document | The link opens in the default web browser after confirmation |

Office Kit has no server of its own, collects no usage data (no telemetry) and has no automatic updates.
Document editing, PDF tools and personal-data masking run entirely on the user's PC.

## Installation and removal / 설치와 제거

The installer installs for the current user only and registers Office Kit as an "Open with" candidate and a context
menu entry for document types; it does not change the default programs. Uninstall from Windows Settings → Apps, which
removes these registrations. Details: [README — 설치](../README.md#설치).
