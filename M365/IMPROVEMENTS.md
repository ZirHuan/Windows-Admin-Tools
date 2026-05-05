# Improvements & Enhancement Proposals

## Get-LicensedUsers.ps1

This document tracks enhancement proposals and improvement ideas for `Get-LicensedUsers.ps1`.

### Overview

`Get-LicensedUsers.ps1` exports detailed licensed user information from M365 tenants, including:
- User identities (UPN, Display Name, Department)
- License SKUs and assignment dates
- Last sign-in activity
- MFA status

Output is generated in both HTML and CSV formats.

### Proposed Enhancements

#### 1. License Lifecycle Tracking
- **Description:** Add fields to track license assignment and expiration dates
- **Benefit:** Helps identify users with upcoming license expirations or recently assigned licenses
- **Implementation:** Query Microsoft Graph for `assignedLicenses` with detailed assignment metadata
- **Priority:** Medium

#### 2. Enhanced MFA Reporting
- **Description:** Expand MFA status to include authentication methods, enrollment status, and enforcement policies
- **Benefit:** Better visibility into MFA coverage and adoption
- **Implementation:** Use `UserAuthenticationMethod.Read.All` scope to pull auth method details
- **Priority:** Medium

#### 3. Last Sign-In Trending
- **Description:** Track sign-in activity over time windows (last 7, 30, 90 days) to identify inactive users
- **Benefit:** Identify dormant accounts for cleanup or relicensing decisions
- **Implementation:** Query `signInActivity` with configurable threshold parameters
- **Priority:** Medium

#### 4. License Conflict Detection
- **Description:** Flag users with conflicting or redundant license assignments
- **Benefit:** Identifies potential cost optimization opportunities
- **Implementation:** Cross-reference assigned SKUs against department/role patterns
- **Priority:** Low

#### 5. Department-Based Insights
- **Description:** Add department-level aggregation and summary statistics
- **Benefit:** Helps analyze licensing spend by organizational unit
- **Implementation:** Group by `department` property and calculate SKU distribution
- **Priority:** Low

#### 6. Custom Filters & Export Options
- **Description:** Add parameters for filtering by license type, MFA status, or inactivity threshold
- **Benefit:** Reduces post-processing and enables targeted reporting
- **Implementation:** Add optional `-LicenseFilter`, `-MFARequired`, `-InactiveThresholdDays` parameters
- **Priority:** Medium

#### 7. Performance Optimization
- **Description:** Implement batch querying and pagination for large tenants
- **Benefit:** Faster execution and reduced API throttling
- **Implementation:** Use Microsoft Graph batch requests; implement proper pagination
- **Priority:** High

### Implementation Status

- [ ] License Lifecycle Tracking
- [ ] Enhanced MFA Reporting
- [ ] Last Sign-In Trending
- [ ] License Conflict Detection
- [ ] Department-Based Insights
- [ ] Custom Filters & Export Options
- [ ] Performance Optimization

---

**Last Updated:** 2026-05-05  
**Repository:** [Windows-Admin-Tools](https://github.com/ZirHuan/Windows-Admin-Tools)
