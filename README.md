# define-monad

項目定義アプリケーション

* Excel VBA とGolangCLI(cobra)で構築する

## golang初期設定

C:/Users/{アカウント}/.cobra.yaml

```yml
author: Teruaki Sato <andrea.pirlo.0529@gmail.com>
license: MIT
useViper: true
```

```console
<!-- goモジュールの作成 -->
go mod init github.com/teru-0529/define-monad
<!-- `cobra-cli`のインストール -->
go install github.com/spf13/cobra-cli@latest
 <!-- プロジェクトのひな型作成 -->
cobra-cli init
<!-- モジュール管理パッケージの整理 -->
go mod tidy
<!-- サブコマンドの登録 -->
cobra-cli add version
cobra-cli add save
cobra-cli add load
```

```console
go install github.com/linyows/git-semv/cmd/git-semv

git semv
git semv -a
git semv now
git semv patch
git semv minor
git semv major
git semv major --pre
git semv major --pre-name rc
git semv patch --bump

git push --delete origin v1.0.0-rc.0

git fetch -p
```

* version サブコマンドの実装

## Excel初期設定

* v2.0.5がそのまま動くことの確認・・・OK
