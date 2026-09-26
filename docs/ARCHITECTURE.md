# Architecture

## 1. Reference

The React/Firebase application in `milance78/interventions-pcd` is the functional reference.

The Excel application must reproduce its user-visible behaviour and business rules where those rules are applicable to a local Excel environment.

## 2. Runtime

The target runtime is a corporate Windows computer with Excel and VBA.

The runtime must not require:

- Node.js
- npm
- React
- Firebase
- a local web server
- browser-based application services

## 3. Update model

During development/distribution, the corporate computer may obtain the newest approved Excel artifact from GitHub.

The update layer will:

1. read the version manifest;
2. compare the local application version with the published version;
3. download the newer Excel artifact when available;
4. verify its checksum;
5. place/open the approved local copy.

The application itself remains a local Excel/VBA application after download.

## 4. Offline mode

The final application must also support a fixed local/offline distribution. The updater is therefore a delivery mechanism, not a runtime dependency.

## 5. Development order

### Phase 1
- workbook shell
- versioning
- update mechanism
- Magic Import

### Phase 2
- Intervention en cours
- field model
- source-specific parsing
- address/client/contact rules

### Phase 3
- CURE
- SNOW
- status/list workflows
- history
- templates

### Phase 4
- export
- security/validation
- regression tests
- final offline package
