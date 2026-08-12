# Q10132 Windows reproduction

This repository contains the final public task package and an independent Windows Server2025 verification workflow for a Node.js HLS publication disposition task.

The workflow checks all four attachment hashes, scans archive members for platform-specific executables, runs the actual Node.js processing chain in two clean directories whose names contain Chinese characters and spaces, compares every generated deliverable with the Reference, exercises a policy-order mutation, and confirms that missing required input fails without leaving output.

Business processing is offline. Network access is used only by GitHub Actions setup steps when obtaining Node.js.
