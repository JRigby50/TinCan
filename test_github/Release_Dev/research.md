Releases

Query
Parameter	Example

tag	        https://github.com/octo-org/octo-repo/releases/new?tag=v1.0.1 creates a release based on a tag named "v1.0.1".
target	    https://github.com/octo-org/octo-repo/releases/new?target=release-1.0.1 creates a release based on the latest commit to the "release-1.0.1" branch.
title	    https://github.com/octo-org/octo-repo/releases/new?tag=v1.0.1&title=octo-1.0.1 creates a release named "octo-1.0.1" based on a tag named "v1.0.1".
body	    https://github.com/octo-org/octo-repo/releases/new?body=Adds+widgets+support creates a release with the description "Adds widget support" in the release body.
prerelease	https://github.com/octo-org/octo-repo/releases/new?prerelease=1 creates a release that will be identified as non-production ready.


Pull Requests

Query
Parameter	Example

quick_pull	https://github.com/octo-org/octo-repo/compare/main...my-branch?quick_pull=1 creates a pull request that compares the base branch main and head branch my-branch. The quick_pull=1 query brings you directly to the "Open a pull request" page.
title	    https://github.com/octo-org/octo-repo/compare/main...my-branch?quick_pull=1&labels=bug&title=Bug+fix creates a pull request with the label "bug" and title "Bug fix."
body	    https://github.com/octo-org/octo-repo/compare/main...my-branch?quick_pull=1&title=Bug+fix&body=Describe+the+fix. creates a pull request with the title "Bug fix" and the   comment "Describe the fix" in the pull request body.
labels	    https://github.com/octo-org/octo-repo/compare/main...my-branch?quick_pull=1&labels=help+wanted,bug creates a pull request with the labels "help wanted" and "bug".
milestone	https://github.com/octo-org/octo-repo/compare/main...my-branch?quick_pull=1&milestone=testing+milestones creates a pull request with the milestone "testing milestones."
assignees	https://github.com/octo-org/octo-repo/compare/main...my-branch?quick_pull=1&assignees=octocat creates a pull request and assigns it to @octocat.
projects	https://github.com/octo-org/octo-repo/compare/main...my-branch?quick_pull=1&title=Bug+fix&projects=octo-org/1 creates a pull request with the title "Bug fix" and adds it to the organization's project 1.
template	https://github.com/octo-org/octo-repo/compare/main...my-branch?quick_pull=1&template=issue_template.md creates a pull request with a template in the pull request body. The template query parameter works with templates stored in a PULL_REQUEST_TEMPLATE subdirectory within the root, docs/ or .github/ directory in a repository.

# GitHub CLI api
# https://cli.github.com/manual/gh_api

gh api \
  --method POST \
  -H "Accept: application/vnd.github+json" \
  -H "X-GitHub-Api-Version: 2022-11-28" \
  --hostname HOSTNAME \
  /repos/OWNER/REPO/pulls \
   -f "title=Amazing new feature" -f "body=Please pull these awesome changes in!" -f "head=octocat:new-feature" -f "base=master"