## [0.14.3](https://github.com/Maxim-Mazurok/teams-api/compare/v0.14.2...v0.14.3) (2026-03-26)


### Bug Fixes

* cache tokens without --email using default account key ([#8](https://github.com/Maxim-Mazurok/teams-api/issues/8)) ([7d17f60](https://github.com/Maxim-Mazurok/teams-api/commit/7d17f60a086f00d0e66ee019dba515ea57ab559d))

## [0.14.2](https://github.com/Maxim-Mazurok/teams-api/compare/v0.14.1...v0.14.2) (2026-03-26)


### Bug Fixes

* clear substrate token polling interval on timeout ([#7](https://github.com/Maxim-Mazurok/teams-api/issues/7)) ([edd3e81](https://github.com/Maxim-Mazurok/teams-api/commit/edd3e819a223d544a38de5e2b3b51dfdbf768631))

## [0.14.1](https://github.com/Maxim-Mazurok/teams-api/compare/v0.14.0...v0.14.1) (2026-03-26)


### Bug Fixes

* cli hangs after auth with installed browser ([#6](https://github.com/Maxim-Mazurok/teams-api/issues/6)) ([d2f9a96](https://github.com/Maxim-Mazurok/teams-api/commit/d2f9a9619f1b27bcdd1419e2ed5e2994584ed1e4))

# [0.14.0](https://github.com/Maxim-Mazurok/teams-api/compare/v0.13.0...v0.14.0) (2026-03-26)


### Features

* add functionality to add and remove reactions on messages ([3f90611](https://github.com/Maxim-Mazurok/teams-api/commit/3f9061117b9973309bedb25a9866a0c5cf6dd6f1))

# [0.13.0](https://github.com/Maxim-Mazurok/teams-api/compare/v0.12.0...v0.13.0) (2026-03-26)


### Features

* cross-platform auth with smart login and secure credential storage ([#5](https://github.com/Maxim-Mazurok/teams-api/issues/5)) ([f7c996a](https://github.com/Maxim-Mazurok/teams-api/commit/f7c996a089d942fffb865f51690d5d3a5c110b48))

# [0.12.0](https://github.com/Maxim-Mazurok/teams-api/compare/v0.11.0...v0.12.0) (2026-03-26)


### Features

* implement file sharing link creation for uploaded files ([80ca06b](https://github.com/Maxim-Mazurok/teams-api/commit/80ca06b96f697ca122ed02c9f246e4ef3d38066e))

# [0.11.0](https://github.com/Maxim-Mazurok/teams-api/compare/v0.10.2...v0.11.0) (2026-03-25)


### Features

* enrich reaction and follower display names from profiles ([571ff47](https://github.com/Maxim-Mazurok/teams-api/commit/571ff47a6a57eec6dac68161580d20e7eca1c0ac))

## [0.10.2](https://github.com/Maxim-Mazurok/teams-api/compare/v0.10.1...v0.10.2) (2026-03-24)


### Bug Fixes

* simplify conversation identity resolution ([71ca82e](https://github.com/Maxim-Mazurok/teams-api/commit/71ca82eacb5ff721e11b8738940a49dea413b942))

## [0.10.1](https://github.com/Maxim-Mazurok/teams-api/compare/v0.10.0...v0.10.1) (2026-03-24)

# [0.10.0](https://github.com/Maxim-Mazurok/teams-api/compare/v0.9.0...v0.10.0) (2026-03-24)


### Features

* add scheduling functionality for messages and enhance file download actions ([2aa2ad2](https://github.com/Maxim-Mazurok/teams-api/commit/2aa2ad28868525d86fec3a200f2ba5b986c81209))

# [0.9.0](https://github.com/Maxim-Mazurok/teams-api/compare/v0.8.0...v0.9.0) (2026-03-24)


### Features

* add file download action and attachment utilities for Teams messages ([8bd751c](https://github.com/Maxim-Mazurok/teams-api/commit/8bd751ccccb87e0927016b004a4b1bd3c6b34605))

# [0.8.0](https://github.com/Maxim-Mazurok/teams-api/compare/v0.7.0...v0.8.0) (2026-03-24)


### Features

* add search and utility actions for people and chats ([937e077](https://github.com/Maxim-Mazurok/teams-api/commit/937e07785e5d05288c92ff47ae9addb74a5d471e))

# [0.7.0](https://github.com/Maxim-Mazurok/teams-api/compare/v0.6.0...v0.7.0) (2026-03-24)


### Features

* replace SKILL.md with MCP server instructions field ([ebd0871](https://github.com/Maxim-Mazurok/teams-api/commit/ebd0871c54f8dce4030980bd28446342e547df87))

# [0.6.0](https://github.com/Maxim-Mazurok/teams-api/compare/v0.5.2...v0.6.0) (2026-03-24)


### Features

* add followers field to messages and update related parsing logic ([8b86c89](https://github.com/Maxim-Mazurok/teams-api/commit/8b86c89c918a46e4181acc2cec7b26d42fda4632))
* implement auto-login, debug session, and interactive login for Teams authentication ([7c25d27](https://github.com/Maxim-Mazurok/teams-api/commit/7c25d2774c3ee98a38a1994a60fbd5e5fb64518c))

## [0.5.2](https://github.com/Maxim-Mazurok/teams-api/compare/v0.5.1...v0.5.2) (2026-03-24)


### Bug Fixes

* update messageId handling to use OriginalArrivalTime for consistency ([3d1681e](https://github.com/Maxim-Mazurok/teams-api/commit/3d1681ed053406fed354e3afa9cf34b1a91d1c63))

## [0.5.1](https://github.com/Maxim-Mazurok/teams-api/compare/v0.5.0...v0.5.1) (2026-03-24)


### Bug Fixes

* throw ApiAuthError for missing bearer token in fetchProfiles and fallback scenarios ([11692dd](https://github.com/Maxim-Mazurok/teams-api/commit/11692dd7ebe69513795cbf5520eefab37df4040e))

# [0.5.0](https://github.com/Maxim-Mazurok/teams-api/compare/v0.4.0...v0.5.0) (2026-03-24)


### Features

* add delete message functionality to API and client ([c46457a](https://github.com/Maxim-Mazurok/teams-api/commit/c46457ab256ed03395479386593dab8f8df55046))

# [0.4.0](https://github.com/Maxim-Mazurok/teams-api/compare/v0.3.3...v0.4.0) (2026-03-24)


### Features

* add edit message functionality to API and client ([640f0f4](https://github.com/Maxim-Mazurok/teams-api/commit/640f0f4268813c18e189b16c7c414b6f015020d8))

## [0.3.3](https://github.com/Maxim-Mazurok/teams-api/compare/v0.3.2...v0.3.3) (2026-03-24)


### Bug Fixes

* throw ApiAuthError when substrate token is missing in search functions ([2ef4539](https://github.com/Maxim-Mazurok/teams-api/commit/2ef45394ad46f32dd081d122f4c28f767a7eea9d))

## [0.3.2](https://github.com/Maxim-Mazurok/teams-api/compare/v0.3.1...v0.3.2) (2026-03-20)

## [0.3.1](https://github.com/Maxim-Mazurok/teams-api/compare/v0.3.0...v0.3.1) (2026-03-19)


### Bug Fixes

* use teams-api-mcp binary in MCP configs instead of CLI binary ([4fd52ea](https://github.com/Maxim-Mazurok/teams-api/commit/4fd52eaf11aa2a5fb72c2675e70ee5a7c6d986f9))

# [0.3.0](https://github.com/Maxim-Mazurok/teams-api/compare/v0.2.7...v0.3.0) (2026-03-19)


### Features

* prompt AI agent for email instead of requiring it in config ([0c908f0](https://github.com/Maxim-Mazurok/teams-api/commit/0c908f0e64d2bf94e364a19b801f3781fd2328f1))

## [0.2.7](https://github.com/Maxim-Mazurok/teams-api/compare/v0.2.6...v0.2.7) (2026-03-19)

## [0.2.6](https://github.com/Maxim-Mazurok/teams-api/compare/v0.2.5...v0.2.6) (2026-03-19)


### Bug Fixes

* dynamically inject version into server.json for MCP Registry publish ([ca34c9f](https://github.com/Maxim-Mazurok/teams-api/commit/ca34c9fb45fea9f3cb4cc6719deda92289cfd410))

## [0.2.5](https://github.com/Maxim-Mazurok/teams-api/compare/v0.2.4...v0.2.5) (2026-03-19)


### Bug Fixes

* update server.json version and remove pinned package version ([6022f6e](https://github.com/Maxim-Mazurok/teams-api/commit/6022f6e4e3d5359891ced38e498c854a80563781))

## [0.2.4](https://github.com/Maxim-Mazurok/teams-api/compare/v0.2.3...v0.2.4) (2026-03-19)


### Bug Fixes

* use correct case for GitHub username in MCP Registry name ([d3eec10](https://github.com/Maxim-Mazurok/teams-api/commit/d3eec10f0db1b204dfa3388f6bca7a18aed351d4))

## [0.2.3](https://github.com/Maxim-Mazurok/teams-api/compare/v0.2.2...v0.2.3) (2026-03-19)


### Bug Fixes

* shorten MCP Registry description to fit 100 char limit ([13f851f](https://github.com/Maxim-Mazurok/teams-api/commit/13f851fce1804e09faf88a4aa55ff519068c69f5))

## [0.2.2](https://github.com/Maxim-Mazurok/teams-api/compare/v0.2.1...v0.2.2) (2026-03-19)


### Bug Fixes

* correct server.json packageArguments to runtimeArguments format ([c6cb744](https://github.com/Maxim-Mazurok/teams-api/commit/c6cb744fc8e2bfca70f3633f286d0baac610cb14))
* fix Prettier formatting and ignore auto-generated CHANGELOG.md ([75a886d](https://github.com/Maxim-Mazurok/teams-api/commit/75a886d97f4244af9afd28d25617007af728fe83))

## [0.2.1](https://github.com/Maxim-Mazurok/teams-api/compare/v0.2.0...v0.2.1) (2026-03-19)

### Bug Fixes

- add repository URL to package.json for npm provenance verification ([24de50d](https://github.com/Maxim-Mazurok/teams-api/commit/24de50d78fd2d00d8a483ac7c8a14164e7be1801))

# [0.2.0](https://github.com/Maxim-Mazurok/teams-api/compare/v0.1.1...v0.2.0) (2026-03-19)

### Features

- rename package to teams-api, add region detection, MCP registry metadata, and provenance publishing ([bf26c99](https://github.com/Maxim-Mazurok/teams-api/commit/bf26c99fe1019cb0f3cdc5c17f479c6d91e9a22e))

## [0.1.1](https://github.com/Maxim-Mazurok/teams-api/compare/v0.1.0...v0.1.1) (2026-03-19)

# Changelog

All notable changes to this project will be documented in this file.
