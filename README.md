# lightsheet
Lightsheet is a lite Excel like JavaScript component for organizes data into spreadsheet elements, with the aim of providing common spreadsheet function such as sorting, filtering, formatting, formulas. Lightsheet aims to be as lightweight as possible and therefore, Lightsheet doesn't use any Js framework like vue or react. 

![G24-ProjPoster-v1-1](https://github.com/lightsheet-team/lightsheet/assets/47510107/c46feaa0-e424-4575-b0d6-7d24c47e3ee0)


## Developer guide
### How to run for developement

- Install node on your computer, version 20+
- Run `npm install `
- Run `npm run dev `

### How to build for production

- Install node on your computer, version 20+
- Run `npm install `
- Run `npm run build `
- Use the content of `dist`, check the folder `pure_js_runner_sample` for an example

### Starting work on a new feature
0. Make sure the feature has an issue. If not, create one.
1. Go to the issue. Click "Create a branch" under "Development".
2. Name the branch something meaningful. Use `kebab-case`. Leave out the issue number.

### Finishing work on a feature
1. Preferably, don't squash your commits. Definitely don't squash the commits if you're not the only one who worked on the feature.
2. Lint your code with `npm run lint`.
3. Make a pull request from the feature branch to `main`.
4. Assign someone to review the PR.
