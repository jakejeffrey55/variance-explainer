/**
 * Development seed.
 *
 * Builds a Dallas–Fort Worth portfolio that exercises every gate the app has to
 * enforce: pending vs. approved vendors, expired compliance, emergency
 * eligibility, budget caps, bid withdrawal, approval flags, contract-required
 * GC jobs, an unclaimed emergency past its response window, and a claimed one
 * with response/arrival timing.
 *
 * Run with: npm run db:seed   (destructive — dev/staging only)
 */
import { PrismaClient, Prisma } from "@prisma/client";
import bcrypt from "bcryptjs";

const prisma = new PrismaClient();

const ADMIN_PASSWORD = "Admin123!";
const VENDOR_PASSWORD = "Vendor123!";

const now = new Date();
const minutes = (n: number) => n * 60 * 1000;
const hours = (n: number) => n * minutes(60);
const days = (n: number) => n * hours(24);
const ago = (ms: number) => new Date(now.getTime() - ms);
const ahead = (ms: number) => new Date(now.getTime() + ms);
const money = (n: number) => new Prisma.Decimal(n.toFixed(2));

export const DEFAULT_TRUST_WEIGHTS = {
  avgJobRating: 0.4,
  complianceUptime: 0.25,
  googleRating: 0.2,
  experience: 0.15,
  // Reliability signals adjust the weighted base rather than adding weight.
  noShowPenaltyPerIncident: 4,
  missedDeadlinePenaltyPerIncident: 2,
  maxReliabilityPenalty: 20,
  emergencyResponseBonusMax: 5,
  emergencyResponseTargetMinutes: 10,
};

async function reset() {
  // Child-first so foreign keys never block the wipe.
  await prisma.$transaction([
    prisma.chatMessage.deleteMany(),
    prisma.chatThread.deleteMany(),
    prisma.notificationLog.deleteMany(),
    prisma.activityLog.deleteMany(),
    prisma.jobDocument.deleteMany(),
    prisma.approvalFlag.deleteMany(),
    prisma.jobRating.deleteMany(),
    prisma.trustScoreSnapshot.deleteMany(),
    prisma.contract.deleteMany(),
    prisma.requisition.deleteMany(),
    prisma.jobInvitation.deleteMany(),
  ]);
  // Jobs point at their awarded bid; clear the pointer before deleting bids.
  await prisma.job.updateMany({ data: { awardedBidId: null } });
  await prisma.$transaction([
    prisma.bid.deleteMany(),
    prisma.job.deleteMany(),
    prisma.vendorAvailability.deleteMany(),
    prisma.vendorUser.deleteMany(),
    prisma.vendor.deleteMany(),
    prisma.rentRoll.deleteMany(),
    prisma.property.deleteMany(),
    prisma.importBatch.deleteMany(),
    prisma.pushSubscription.deleteMany(),
    prisma.approvalSettings.deleteMany(),
    prisma.adminUser.deleteMany(),
  ]);
}

async function main() {
  console.log("Resetting database…");
  await reset();

  const adminHash = await bcrypt.hash(ADMIN_PASSWORD, 10);
  const vendorHash = await bcrypt.hash(VENDOR_PASSWORD, 10);

  // -------------------------------------------------------------------------
  // Admins (property management staff)
  // -------------------------------------------------------------------------
  const owner = await prisma.adminUser.create({
    data: {
      email: "dana.reyes@cortland-example.com",
      passwordHash: adminHash,
      name: "Dana Reyes",
      phone: "+12145550101",
      role: "OWNER",
    },
  });
  const manager = await prisma.adminUser.create({
    data: {
      email: "marcus.hale@cortland-example.com",
      passwordHash: adminHash,
      name: "Marcus Hale",
      phone: "+12145550102",
      role: "MANAGER",
    },
  });
  const staff = await prisma.adminUser.create({
    data: {
      email: "priya.nair@cortland-example.com",
      passwordHash: adminHash,
      name: "Priya Nair",
      phone: "+12145550103",
      role: "STAFF",
    },
  });

  await prisma.approvalSettings.create({
    data: {
      id: "default",
      aboveAverageThresholdPct: 7.5,
      contractThresholdAmount: money(5000),
      emergencyResponseMinutes: 15,
      emergencyRadiusExpansionMiles: 25,
      defaultServiceRadiusMiles: 25,
      complianceExpiryWarningDays: 30,
      trustScoreWeights: DEFAULT_TRUST_WEIGHTS,
      updatedByAdminId: owner.id,
    },
  });

  // -------------------------------------------------------------------------
  // Properties — as if imported from a Power BI CSV export
  // -------------------------------------------------------------------------
  const csvBatch = await prisma.importBatch.create({
    data: {
      kind: "PROPERTY",
      providerKey: "csv",
      status: "COMPLETED",
      fileName: "powerbi_property_export_2026_08.csv",
      rowsTotal: 5,
      rowsImported: 5,
      adminUserId: owner.id,
      startedAt: ago(days(9)),
      finishedAt: ago(days(9) - minutes(2)),
    },
  });

  const propertyRows = [
    {
      externalId: "CRT-DAL-0148",
      name: "Cortland Uptown Dallas",
      addressLine1: "2801 Cedar Springs Rd",
      city: "Dallas",
      state: "TX",
      postalCode: "75201",
      latitude: 32.7997,
      longitude: -96.8065,
      unitCount: 312,
      propertyManagerName: "Alicia Gomez",
      propertyManagerEmail: "alicia.gomez@cortland-example.com",
      propertyManagerPhone: "+12145550201",
    },
    {
      externalId: "CRT-PLN-0203",
      name: "Cortland Legacy Plano",
      addressLine1: "5800 Headquarters Dr",
      city: "Plano",
      state: "TX",
      postalCode: "75024",
      latitude: 33.0793,
      longitude: -96.8214,
      unitCount: 268,
      propertyManagerName: "Ben Okafor",
      propertyManagerEmail: "ben.okafor@cortland-example.com",
      propertyManagerPhone: "+19725550202",
    },
    {
      externalId: "CRT-FTW-0091",
      name: "Cortland Riverbend Fort Worth",
      addressLine1: "1400 Rogers Rd",
      city: "Fort Worth",
      state: "TX",
      postalCode: "76107",
      latitude: 32.7381,
      longitude: -97.3639,
      unitCount: 184,
      propertyManagerName: "Kelly Underwood",
      propertyManagerEmail: "kelly.underwood@cortland-example.com",
      propertyManagerPhone: "+18175550203",
    },
    {
      externalId: "CRT-ARL-0117",
      name: "Cortland Arlington Commons",
      addressLine1: "2100 E Lamar Blvd",
      city: "Arlington",
      state: "TX",
      postalCode: "76006",
      latitude: 32.7554,
      longitude: -97.0839,
      unitCount: 240,
      propertyManagerName: "Rosa Delgado",
      propertyManagerEmail: "rosa.delgado@cortland-example.com",
      propertyManagerPhone: "+18175550204",
    },
    {
      externalId: "CRT-FRS-0325",
      name: "Cortland Frisco Station",
      addressLine1: "6250 Frisco Square Blvd",
      city: "Frisco",
      state: "TX",
      postalCode: "75034",
      latitude: 33.1541,
      longitude: -96.8236,
      unitCount: 196,
      propertyManagerName: "Trent Wu",
      propertyManagerEmail: "trent.wu@cortland-example.com",
      propertyManagerPhone: "+19725550205",
    },
  ];

  const [uptown, legacy, riverbend, arlington, frisco] = await Promise.all(
    propertyRows.map((p) =>
      prisma.property.create({
        data: { ...p, source: "CSV", sourceBatchId: csvBatch.id, lastSyncedAt: ago(days(9)) },
      }),
    ),
  );

  // Rent roll (optional import; only informs due-date suggestions)
  await prisma.rentRoll.createMany({
    data: [
      {
        propertyId: uptown.id,
        unitNumber: "214",
        unitType: "A2 / 1x1",
        squareFeet: 742,
        marketRent: money(1795),
        currentRent: money(1720),
        status: "Notice",
        leaseEndDate: ahead(days(11)),
        moveOutDate: ahead(days(11)),
        moveInDate: ahead(days(26)),
      },
      {
        propertyId: uptown.id,
        unitNumber: "508",
        unitType: "B1 / 2x2",
        squareFeet: 1105,
        marketRent: money(2450),
        currentRent: money(2380),
        status: "Vacant Ready",
        moveOutDate: ago(days(6)),
      },
      {
        propertyId: legacy.id,
        unitNumber: "1122",
        unitType: "A1 / 1x1",
        squareFeet: 688,
        marketRent: money(1620),
        status: "Occupied",
        leaseEndDate: ahead(days(84)),
      },
      {
        propertyId: frisco.id,
        unitNumber: "306",
        unitType: "B2 / 2x2",
        squareFeet: 1180,
        marketRent: money(2610),
        status: "Vacant Unrented",
        moveOutDate: ago(days(19)),
      },
    ],
  });

  // -------------------------------------------------------------------------
  // Vendors — every access state the gates have to handle
  // -------------------------------------------------------------------------
  type VendorSeed = {
    key: string;
    data: Prisma.VendorCreateInput;
    user: { email: string; name: string; phone: string };
  };

  const vendorSeeds: VendorSeed[] = [
    {
      key: "loneStar",
      data: {
        companyName: "Lone Star Make-Ready Co.",
        contactName: "Rick Alvarez",
        email: "ops@lonestarmakeready.example.com",
        phone: "+12145551001",
        addressLine1: "1120 Empire Central",
        city: "Dallas",
        state: "TX",
        postalCode: "75247",
        latitude: 32.8207,
        longitude: -96.8719,
        serviceRadiusMiles: 30,
        serviceCategories: ["MAKE_READY", "PAINTING", "FLOORING", "CLEANING", "DRYWALL"],
        accountStatus: "ACTIVE",
        emergencyEligible: true,
        complianceStatus: "COMPLIANT",
        vendorplyId: "VP-100482",
        complianceExpiresAt: ahead(days(214)),
        insuranceExpiresAt: ahead(days(214)),
        licenseNumber: "TX-MR-88214",
        w9OnFile: true,
        lastCredentialSyncAt: ago(days(3)),
        googlePlaceId: "ChIJseed_lonestar",
        googleRating: 4.6,
        googleRatingCount: 187,
        googleRatingFetchedAt: ago(days(2)),
        yearsInBusiness: 12,
        approvedAt: ago(days(400)),
        approvedBy: { connect: { id: owner.id } },
      },
      user: { email: "rick@lonestarmakeready.example.com", name: "Rick Alvarez", phone: "+12145551001" },
    },
    {
      key: "metroplexHvac",
      data: {
        companyName: "Metroplex HVAC & Air",
        contactName: "Sandra Pham",
        email: "dispatch@metroplexhvac.example.com",
        phone: "+19725551002",
        addressLine1: "3400 Preston Rd",
        city: "Plano",
        state: "TX",
        postalCode: "75093",
        latitude: 33.0335,
        longitude: -96.8021,
        serviceRadiusMiles: 40,
        serviceCategories: ["HVAC", "APPLIANCE", "ELECTRICAL"],
        accountStatus: "ACTIVE",
        emergencyEligible: true,
        complianceStatus: "COMPLIANT",
        vendorplyId: "VP-100915",
        complianceExpiresAt: ahead(days(96)),
        insuranceExpiresAt: ahead(days(96)),
        licenseNumber: "TACLA-42117C",
        w9OnFile: true,
        lastCredentialSyncAt: ago(days(3)),
        googlePlaceId: "ChIJseed_metroplex",
        googleRating: 4.8,
        googleRatingCount: 512,
        googleRatingFetchedAt: ago(days(2)),
        yearsInBusiness: 19,
        approvedAt: ago(days(365)),
        approvedBy: { connect: { id: owner.id } },
      },
      user: { email: "sandra@metroplexhvac.example.com", name: "Sandra Pham", phone: "+19725551002" },
    },
    {
      key: "rapidDry",
      data: {
        companyName: "RapidDry Water Restoration",
        contactName: "Terrence Boyd",
        email: "24hr@rapiddry.example.com",
        phone: "+18175551003",
        addressLine1: "905 Avenue H E",
        city: "Arlington",
        state: "TX",
        postalCode: "76011",
        latitude: 32.7602,
        longitude: -97.0783,
        serviceRadiusMiles: 50,
        serviceCategories: ["WATER_MITIGATION", "PLUMBING", "MAKE_READY"],
        accountStatus: "ACTIVE",
        emergencyEligible: true,
        complianceStatus: "COMPLIANT",
        vendorplyId: "VP-101330",
        complianceExpiresAt: ahead(days(150)),
        insuranceExpiresAt: ahead(days(150)),
        licenseNumber: "IICRC-77120",
        w9OnFile: true,
        lastCredentialSyncAt: ago(days(3)),
        googlePlaceId: "ChIJseed_rapiddry",
        googleRating: 4.4,
        googleRatingCount: 96,
        googleRatingFetchedAt: ago(days(2)),
        yearsInBusiness: 8,
        approvedAt: ago(days(220)),
        approvedBy: { connect: { id: manager.id } },
      },
      user: { email: "terrence@rapiddry.example.com", name: "Terrence Boyd", phone: "+18175551003" },
    },
    {
      key: "bluebonnet",
      data: {
        companyName: "Bluebonnet General Contracting",
        contactName: "Marisol Vega",
        email: "bids@bluebonnetgc.example.com",
        phone: "+18175551004",
        addressLine1: "2400 W 7th St",
        city: "Fort Worth",
        state: "TX",
        postalCode: "76107",
        latitude: 32.7495,
        longitude: -97.3585,
        serviceRadiusMiles: 45,
        serviceCategories: ["GENERAL_CONTRACTING", "DRYWALL", "ROOFING", "PAINTING"],
        accountStatus: "ACTIVE",
        emergencyEligible: false,
        complianceStatus: "COMPLIANT",
        vendorplyId: "VP-100277",
        complianceExpiresAt: ahead(days(310)),
        insuranceExpiresAt: ahead(days(310)),
        licenseNumber: "TX-GC-55901",
        w9OnFile: true,
        lastCredentialSyncAt: ago(days(3)),
        googlePlaceId: "ChIJseed_bluebonnet",
        googleRating: 4.2,
        googleRatingCount: 64,
        googleRatingFetchedAt: ago(days(2)),
        yearsInBusiness: 22,
        approvedAt: ago(days(500)),
        approvedBy: { connect: { id: owner.id } },
      },
      user: { email: "marisol@bluebonnetgc.example.com", name: "Marisol Vega", phone: "+18175551004" },
    },
    {
      key: "trinityFlooring",
      data: {
        companyName: "Trinity Flooring Group",
        contactName: "Owen Baptiste",
        email: "schedule@trinityflooring.example.com",
        phone: "+19725551005",
        addressLine1: "8500 Gaylord Pkwy",
        city: "Frisco",
        state: "TX",
        postalCode: "75034",
        latitude: 33.0946,
        longitude: -96.8206,
        serviceRadiusMiles: 25,
        serviceCategories: ["FLOORING", "MAKE_READY"],
        accountStatus: "ACTIVE",
        emergencyEligible: false,
        // Expiring inside the 30-day warning window — shows on the admin dashboard.
        complianceStatus: "EXPIRING_SOON",
        vendorplyId: "VP-102044",
        complianceExpiresAt: ahead(days(18)),
        insuranceExpiresAt: ahead(days(18)),
        w9OnFile: true,
        lastCredentialSyncAt: ago(days(3)),
        googleRating: 4.1,
        googleRatingCount: 41,
        googleRatingFetchedAt: ago(days(2)),
        yearsInBusiness: 6,
        approvedAt: ago(days(140)),
        approvedBy: { connect: { id: manager.id } },
      },
      user: { email: "owen@trinityflooring.example.com", name: "Owen Baptiste", phone: "+19725551005" },
    },
    {
      key: "sunbeltElectric",
      data: {
        companyName: "Sunbelt Electrical Services",
        contactName: "Dee Whitfield",
        email: "office@sunbeltelectric.example.com",
        phone: "+12145551006",
        addressLine1: "4455 Sigma Rd",
        city: "Dallas",
        state: "TX",
        postalCode: "75244",
        latitude: 32.9256,
        longitude: -96.8256,
        serviceRadiusMiles: 35,
        serviceCategories: ["ELECTRICAL", "APPLIANCE"],
        accountStatus: "ACTIVE",
        // Emergency-eligible but compliance is expired: the compliance gate
        // still blocks bidding and claiming. The two toggles are independent.
        emergencyEligible: true,
        complianceStatus: "EXPIRED",
        vendorplyId: "VP-100777",
        complianceExpiresAt: ago(days(12)),
        insuranceExpiresAt: ago(days(12)),
        licenseNumber: "TECL-31855",
        w9OnFile: true,
        lastCredentialSyncAt: ago(days(3)),
        googleRating: 3.9,
        googleRatingCount: 78,
        googleRatingFetchedAt: ago(days(2)),
        yearsInBusiness: 15,
        approvedAt: ago(days(300)),
        approvedBy: { connect: { id: owner.id } },
      },
      user: { email: "dee@sunbeltelectric.example.com", name: "Dee Whitfield", phone: "+12145551006" },
    },
    {
      key: "ridgeline",
      data: {
        companyName: "Ridgeline Renovations",
        contactName: "Hugo Marin",
        email: "info@ridgelinereno.example.com",
        phone: "+17135551007",
        addressLine1: "1200 Smith St",
        city: "Houston",
        state: "TX",
        postalCode: "77002",
        // ~225 miles from DFW: in good standing, but outside every property's
        // radius — proves the haversine filter, not the status filter, excludes it.
        latitude: 29.7589,
        longitude: -95.3677,
        serviceRadiusMiles: 40,
        serviceCategories: ["MAKE_READY", "GENERAL_CONTRACTING", "PAINTING"],
        accountStatus: "ACTIVE",
        emergencyEligible: true,
        complianceStatus: "COMPLIANT",
        vendorplyId: "VP-103611",
        complianceExpiresAt: ahead(days(260)),
        insuranceExpiresAt: ahead(days(260)),
        w9OnFile: true,
        googleRating: 4.5,
        googleRatingCount: 122,
        yearsInBusiness: 9,
        approvedAt: ago(days(90)),
        approvedBy: { connect: { id: manager.id } },
      },
      user: { email: "hugo@ridgelinereno.example.com", name: "Hugo Marin", phone: "+17135551007" },
    },
    {
      key: "pinnacle",
      data: {
        companyName: "Pinnacle Property Services",
        contactName: "Ava Lindstrom",
        email: "hello@pinnacleprops.example.com",
        phone: "+19725551008",
        addressLine1: "2701 Dallas Pkwy",
        city: "Plano",
        state: "TX",
        postalCode: "75093",
        latitude: 33.0198,
        longitude: -96.8352,
        serviceRadiusMiles: 25,
        serviceCategories: ["MAKE_READY", "CLEANING", "LANDSCAPING"],
        // Brand-new signup: can sign in, sees a pending screen, cannot see or
        // bid on a single job until an admin approves.
        accountStatus: "PENDING",
        emergencyEligible: false,
        complianceStatus: "NOT_SUBMITTED",
        yearsInBusiness: 3,
      },
      user: { email: "ava@pinnacleprops.example.com", name: "Ava Lindstrom", phone: "+19725551008" },
    },
    {
      key: "gulfCoast",
      data: {
        companyName: "Gulf Coast Handyman LLC",
        contactName: "Neil Prather",
        email: "neil@gulfcoasthandyman.example.com",
        phone: "+12145551009",
        addressLine1: "77 Industrial Blvd",
        city: "Dallas",
        state: "TX",
        postalCode: "75207",
        latitude: 32.7801,
        longitude: -96.8299,
        serviceRadiusMiles: 20,
        serviceCategories: ["MAKE_READY", "APPLIANCE"],
        // Suspended: cannot even establish a session.
        accountStatus: "SUSPENDED",
        emergencyEligible: false,
        complianceStatus: "EXPIRED",
        complianceExpiresAt: ago(days(75)),
        internalNotes: "Suspended 2026-06 — repeated no-shows on make-ready turns.",
        yearsInBusiness: 4,
      },
      user: { email: "neil@gulfcoasthandyman.example.com", name: "Neil Prather", phone: "+12145551009" },
    },
  ];

  const vendors: Record<string, { id: string }> = {};
  for (const seed of vendorSeeds) {
    const vendor = await prisma.vendor.create({ data: seed.data });
    vendors[seed.key] = vendor;
    await prisma.vendorUser.create({
      data: {
        vendorId: vendor.id,
        email: seed.user.email,
        passwordHash: vendorHash,
        name: seed.user.name,
        phone: seed.user.phone,
        isPrimary: true,
      },
    });
  }

  // A second login on one vendor account — proves per-vendor (not per-user)
  // data scoping.
  await prisma.vendorUser.create({
    data: {
      vendorId: vendors.loneStar.id,
      email: "coordinator@lonestarmakeready.example.com",
      passwordHash: vendorHash,
      name: "Jo Tran",
      phone: "+12145551011",
    },
  });

  // Blackout dates
  await prisma.vendorAvailability.createMany({
    data: [
      {
        vendorId: vendors.loneStar.id,
        type: "BLACKOUT",
        startDate: ahead(days(6)),
        endDate: ahead(days(13)),
        reason: "Crew on scheduled leave",
      },
      {
        vendorId: vendors.trinityFlooring.id,
        type: "BLACKOUT",
        startDate: ahead(days(1)),
        endDate: ahead(days(4)),
        reason: "Booked on external project",
      },
      {
        vendorId: vendors.metroplexHvac.id,
        type: "REDUCED_CAPACITY",
        startDate: ahead(days(20)),
        endDate: ahead(days(24)),
        reason: "Two techs in certification training",
      },
    ],
  });

  // -------------------------------------------------------------------------
  // Jobs
  // -------------------------------------------------------------------------

  // J-1001 — open make-ready with a hard budget cap and live bidding
  const job1001 = await prisma.job.create({
    data: {
      jobNumber: "J-1001",
      propertyId: uptown.id,
      unitNumber: "214",
      title: "Full make-ready turn — Unit 214",
      description:
        "Paint throughout, replace LVP in living room, deep clean, punch appliances. Unit vacates in 11 days; must be rent-ready before the scheduled move-in.",
      category: "MAKE_READY",
      status: "OPEN",
      budgetMin: money(1800),
      budgetMax: money(3200),
      enforceBudgetCap: true,
      bidDeadline: ahead(days(5)),
      dueDate: ahead(days(24)),
      createdByAdminId: staff.id,
      createdAt: ago(days(2)),
    },
  });

  // J-1002 — GC scope over the contract threshold, bids in, awaiting approval
  const job1002 = await prisma.job.create({
    data: {
      jobNumber: "J-1002",
      propertyId: riverbend.id,
      unitNumber: "Clubhouse",
      title: "Clubhouse restroom remodel",
      description:
        "Demo and rebuild two clubhouse restrooms: tile, fixtures, partitions, ADA compliance review.",
      category: "GENERAL_CONTRACTING",
      status: "AWAITING_APPROVAL",
      budgetMin: money(6000),
      budgetMax: money(12000),
      enforceBudgetCap: false,
      bidDeadline: ago(hours(6)),
      dueDate: ahead(days(45)),
      createdByAdminId: manager.id,
      createdAt: ago(days(12)),
    },
  });

  // J-1003 — EMERGENCY, dispatched 41 minutes ago, still unclaimed and escalated
  const job1003 = await prisma.job.create({
    data: {
      jobNumber: "J-1003",
      propertyId: legacy.id,
      unitNumber: "1122",
      title: "EMERGENCY — No cooling, occupied unit",
      description:
        "Resident reports AC blowing warm since this morning. Indoor temp 88°F, elderly resident on site.",
      category: "HVAC",
      status: "OPEN",
      priority: "EMERGENCY",
      emergencyCategory: "AC_HVAC",
      budgetMin: money(250),
      budgetMax: money(1500),
      enforceBudgetCap: false,
      responseDeadlineMinutes: 15,
      dispatchedAt: ago(minutes(41)),
      escalatedAt: ago(minutes(26)),
      escalationRadiusMiles: 65,
      createdByAdminId: staff.id,
      createdAt: ago(minutes(41)),
    },
  });

  // J-1004 — EMERGENCY, claimed in 6 minutes, vendor on site at 32 minutes
  const job1004 = await prisma.job.create({
    data: {
      jobNumber: "J-1004",
      propertyId: arlington.id,
      unitNumber: "308",
      title: "EMERGENCY — Water intrusion from unit above",
      description:
        "Supply line failure in 408 flooding 308. Water shut off at riser; extraction and dry-out needed immediately.",
      category: "WATER_MITIGATION",
      status: "AWARDED",
      priority: "EMERGENCY",
      emergencyCategory: "WATER_EXTRACTION",
      budgetMin: money(800),
      budgetMax: money(4500),
      responseDeadlineMinutes: 15,
      dispatchedAt: ago(hours(20)),
      claimedByVendorId: vendors.rapidDry.id,
      claimedAt: ago(hours(20) - minutes(6)),
      onSiteAt: ago(hours(20) - minutes(32)),
      awardedVendorId: vendors.rapidDry.id,
      awardedAt: ago(hours(20) - minutes(6)),
      startedAt: ago(hours(20) - minutes(32)),
      createdByAdminId: staff.id,
      createdAt: ago(hours(20)),
    },
  });

  // J-1005 — GC roof repair, awarded above the average-bid threshold, contract out for signature
  const job1005 = await prisma.job.create({
    data: {
      jobNumber: "J-1005",
      propertyId: frisco.id,
      unitNumber: "Building C",
      title: "Building C roof section repair",
      description:
        "Replace ~18 squares of damaged shingles, repair decking, reflash two penetrations after hail damage.",
      category: "GENERAL_CONTRACTING",
      status: "AWARDED",
      budgetMin: money(6500),
      budgetMax: money(9500),
      enforceBudgetCap: false,
      bidDeadline: ago(days(4)),
      scheduledStart: ahead(days(3)),
      dueDate: ahead(days(17)),
      awardedVendorId: vendors.bluebonnet.id,
      awardedAt: ago(days(2)),
      createdByAdminId: manager.id,
      createdAt: ago(days(14)),
    },
  });

  // J-1006 — completed turn, rated, approved over budget_max with the cap off
  const job1006 = await prisma.job.create({
    data: {
      jobNumber: "J-1006",
      propertyId: uptown.id,
      unitNumber: "508",
      title: "Make-ready turn — Unit 508",
      description: "Two-tone paint, carpet replacement in both bedrooms, full clean, blind replacement.",
      category: "MAKE_READY",
      status: "COMPLETED",
      budgetMin: money(1500),
      budgetMax: money(2400),
      enforceBudgetCap: false,
      bidDeadline: ago(days(20)),
      awardedVendorId: vendors.loneStar.id,
      awardedAt: ago(days(19)),
      startedAt: ago(days(16)),
      completedAt: ago(days(9)),
      dueDate: ago(days(8)),
      createdByAdminId: staff.id,
      createdAt: ago(days(24)),
    },
  });

  // J-1007 — invite-only GC job
  const job1007 = await prisma.job.create({
    data: {
      jobNumber: "J-1007",
      propertyId: riverbend.id,
      unitNumber: "Breezeway A",
      title: "Stair tread and railing replacement — Breezeway A",
      description: "Replace 22 stair treads and 40 LF of railing to code. Invited GCs only.",
      category: "GENERAL_CONTRACTING",
      status: "OPEN",
      budgetMin: money(4200),
      budgetMax: money(7800),
      enforceBudgetCap: false,
      inviteOnly: true,
      bidDeadline: ahead(days(8)),
      createdByAdminId: manager.id,
      createdAt: ago(days(1)),
    },
  });

  // J-1008 — bidding window already closed
  const job1008 = await prisma.job.create({
    data: {
      jobNumber: "J-1008",
      propertyId: arlington.id,
      unitNumber: "112",
      title: "Dishwasher replacement — Unit 112",
      description: "Remove and replace dishwasher, verify supply and drain connections.",
      category: "APPLIANCE",
      status: "BIDDING_CLOSED",
      budgetMin: money(400),
      budgetMax: money(750),
      enforceBudgetCap: true,
      bidDeadline: ago(days(1)),
      createdByAdminId: staff.id,
      createdAt: ago(days(6)),
    },
  });

  await prisma.jobInvitation.create({
    data: {
      jobId: job1007.id,
      vendorId: vendors.bluebonnet.id,
      invitedByAdminId: manager.id,
      invitedAt: ago(days(1)),
      viewedAt: ago(hours(20)),
    },
  });

  // -------------------------------------------------------------------------
  // Bids
  // -------------------------------------------------------------------------

  // J-1001: two live bids plus one the vendor withdrew (kept in history)
  const bid1001LoneStar = await prisma.bid.create({
    data: {
      jobId: job1001.id,
      vendorId: vendors.loneStar.id,
      amount: money(2680),
      laborCost: money(1780),
      materialCost: money(900),
      status: "SUBMITTED",
      notes: "Can start the day after vacate; 4-day turn including paint cure.",
      estimatedStartDate: ahead(days(12)),
      estimatedCompletionDate: ahead(days(16)),
      submittedAt: ago(days(1)),
    },
  });
  await prisma.bid.create({
    data: {
      jobId: job1001.id,
      vendorId: vendors.trinityFlooring.id,
      amount: money(3050),
      status: "SUBMITTED",
      notes: "LVP in stock; paint subbed to our regular crew.",
      estimatedStartDate: ahead(days(13)),
      estimatedCompletionDate: ahead(days(18)),
      submittedAt: ago(hours(30)),
    },
  });
  await prisma.bid.create({
    data: {
      jobId: job1001.id,
      vendorId: vendors.rapidDry.id,
      amount: money(2900),
      status: "WITHDRAWN",
      notes: "Withdrawing — crew committed to a loss job through month end.",
      submittedAt: ago(days(2)),
      withdrawnAt: ago(hours(20)),
    },
  });

  // J-1002: awaiting approval — this is what the comparison table renders from
  await prisma.bid.create({
    data: {
      jobId: job1002.id,
      vendorId: vendors.bluebonnet.id,
      amount: money(9450),
      laborCost: money(6100),
      materialCost: money(3350),
      status: "SUBMITTED",
      notes: "Includes ADA grab bars and partition replacement; 3-week schedule.",
      estimatedStartDate: ahead(days(9)),
      estimatedCompletionDate: ahead(days(30)),
      submittedAt: ago(days(3)),
    },
  });
  await prisma.bid.create({
    data: {
      jobId: job1002.id,
      vendorId: vendors.ridgeline.id,
      amount: money(8200),
      status: "SUBMITTED",
      notes: "Travel from Houston included. Tile allowance $9/sf.",
      estimatedStartDate: ahead(days(14)),
      estimatedCompletionDate: ahead(days(38)),
      submittedAt: ago(days(2)),
    },
  });

  // J-1005: approved bid sits 12.5% above the three-bid average → flagged
  const bid1005Bluebonnet = await prisma.bid.create({
    data: {
      jobId: job1005.id,
      vendorId: vendors.bluebonnet.id,
      amount: money(8400),
      laborCost: money(5200),
      materialCost: money(3200),
      status: "APPROVED",
      notes: "Crew available immediately; 5-year workmanship warranty.",
      estimatedStartDate: ahead(days(3)),
      estimatedCompletionDate: ahead(days(12)),
      submittedAt: ago(days(6)),
      approvedAt: ago(days(2)),
      approvedByAdminId: manager.id,
    },
  });
  await prisma.bid.create({
    data: {
      jobId: job1005.id,
      vendorId: vendors.ridgeline.id,
      amount: money(6900),
      status: "REJECTED",
      submittedAt: ago(days(7)),
      rejectedAt: ago(days(2)),
      rejectionReason: "Schedule could not meet the pre-inspection date.",
    },
  });
  await prisma.bid.create({
    data: {
      jobId: job1005.id,
      vendorId: vendors.loneStar.id,
      amount: money(7100),
      status: "REJECTED",
      submittedAt: ago(days(6)),
      rejectedAt: ago(days(2)),
      rejectionReason: "Roofing outside of awarded scope of work.",
    },
  });

  // J-1006: completed, approved $250 over budget_max with the cap off → flagged
  const bid1006LoneStar = await prisma.bid.create({
    data: {
      jobId: job1006.id,
      vendorId: vendors.loneStar.id,
      amount: money(2650),
      status: "APPROVED",
      aboveBudgetMax: true,
      notes: "Carpet pricing up since last turn; includes blind replacement.",
      submittedAt: ago(days(21)),
      approvedAt: ago(days(19)),
      approvedByAdminId: staff.id,
    },
  });

  await prisma.job.update({
    where: { id: job1005.id },
    data: { awardedBidId: bid1005Bluebonnet.id },
  });
  await prisma.job.update({
    where: { id: job1006.id },
    data: { awardedBidId: bid1006LoneStar.id },
  });

  // -------------------------------------------------------------------------
  // Approval flags (never block approval — they record it)
  // -------------------------------------------------------------------------
  await prisma.approvalFlag.create({
    data: {
      jobId: job1005.id,
      bidId: bid1005Bluebonnet.id,
      propertyId: frisco.id,
      type: "ABOVE_AVERAGE_THRESHOLD",
      thresholdPct: 7.5,
      averageBidAmount: money(7466.67),
      approvedAmount: money(8400),
      deltaPct: 12.5,
      bidCount: 3,
      note: "Approved on schedule certainty ahead of the pre-inspection date.",
      createdAt: ago(days(2)),
    },
  });
  await prisma.approvalFlag.create({
    data: {
      jobId: job1006.id,
      bidId: bid1006LoneStar.id,
      propertyId: uptown.id,
      type: "ABOVE_BUDGET_CAP",
      approvedAmount: money(2650),
      budgetMax: money(2400),
      deltaPct: 10.42,
      bidCount: 1,
      note: "Sole bidder; carpet material increase accepted.",
      createdAt: ago(days(19)),
      acknowledgedByAdminId: owner.id,
      acknowledgedAt: ago(days(18)),
    },
  });

  // -------------------------------------------------------------------------
  // Requisitions + contracts (mock providers)
  // -------------------------------------------------------------------------
  await prisma.requisition.create({
    data: {
      jobId: job1005.id,
      bidId: bid1005Bluebonnet.id,
      vendorId: vendors.bluebonnet.id,
      providerKey: "mock",
      externalId: "MOCK-REQ-4471",
      status: "ACKNOWLEDGED",
      amount: money(8400),
      requestPayload: { jobNumber: "J-1005", glCode: "5210-ROOF", propertyExternalId: "CRT-FRS-0325" },
      responseBody: { requisitionNumber: "MOCK-REQ-4471", state: "pending_approval" },
      submittedAt: ago(days(2)),
      acknowledgedAt: ago(days(2) - minutes(3)),
    },
  });
  await prisma.contract.create({
    data: {
      jobId: job1005.id,
      bidId: bid1005Bluebonnet.id,
      vendorId: vendors.bluebonnet.id,
      providerKey: "mock",
      externalId: "MOCK-CON-0088",
      status: "PENDING_SIGNATURE",
      amount: money(8400),
      documentUrl: "https://contracts.example.com/mock/MOCK-CON-0088",
      sentAt: ago(days(2) - minutes(5)),
      expiresAt: ahead(days(12)),
      payload: { template: "gc_standard_v3", thresholdAmount: 5000 },
    },
  });
  await prisma.requisition.create({
    data: {
      jobId: job1004.id,
      vendorId: vendors.rapidDry.id,
      providerKey: "mock",
      externalId: "MOCK-REQ-4468",
      status: "ACKNOWLEDGED",
      amount: money(4500),
      requestPayload: { jobNumber: "J-1004", emergency: true, notToExceed: 4500 },
      responseBody: { requisitionNumber: "MOCK-REQ-4468", state: "auto_approved_emergency" },
      submittedAt: ago(hours(20) - minutes(6)),
      acknowledgedAt: ago(hours(20) - minutes(7)),
    },
  });
  await prisma.requisition.create({
    data: {
      jobId: job1006.id,
      bidId: bid1006LoneStar.id,
      vendorId: vendors.loneStar.id,
      providerKey: "mock",
      externalId: "MOCK-REQ-4402",
      status: "ACKNOWLEDGED",
      amount: money(2650),
      submittedAt: ago(days(19)),
      acknowledgedAt: ago(days(19)),
    },
  });

  // -------------------------------------------------------------------------
  // Rating + trust score snapshots
  // -------------------------------------------------------------------------
  await prisma.jobRating.create({
    data: {
      jobId: job1006.id,
      vendorId: vendors.loneStar.id,
      adminUserId: staff.id,
      score: 5,
      noShow: false,
      missedDeadline: false,
      qualityScore: 5,
      communicationScore: 5,
      timelinessScore: 4,
      comment: "Turn delivered a day early. Punch list came back clean on the first walk.",
      createdAt: ago(days(9)),
    },
  });

  const snapshots = [
    {
      vendorId: vendors.loneStar.id,
      score: 91.4,
      avgJobRating: 4.8,
      complianceUptimePct: 100,
      googleRating: 4.6,
      experienceScore: 80,
      reliabilityPenalty: 0,
      emergencyResponseScore: 3.5,
      completedJobCount: 34,
      avgEmergencyResponseMins: 9.2,
    },
    {
      vendorId: vendors.metroplexHvac.id,
      score: 94.1,
      avgJobRating: 4.9,
      complianceUptimePct: 100,
      googleRating: 4.8,
      experienceScore: 95,
      reliabilityPenalty: 0,
      emergencyResponseScore: 5,
      completedJobCount: 58,
      avgEmergencyResponseMins: 6.4,
    },
    {
      vendorId: vendors.rapidDry.id,
      score: 86.7,
      avgJobRating: 4.5,
      complianceUptimePct: 96,
      googleRating: 4.4,
      experienceScore: 60,
      reliabilityPenalty: 2,
      emergencyResponseScore: 4.6,
      noShowCount: 0,
      missedDeadlineCount: 1,
      completedJobCount: 21,
      avgEmergencyResponseMins: 7.8,
    },
    {
      vendorId: vendors.bluebonnet.id,
      score: 79.3,
      avgJobRating: 4.1,
      complianceUptimePct: 98,
      googleRating: 4.2,
      experienceScore: 100,
      reliabilityPenalty: 6,
      noShowCount: 1,
      missedDeadlineCount: 1,
      completedJobCount: 47,
    },
  ];
  for (const snap of snapshots) {
    await prisma.trustScoreSnapshot.create({
      data: {
        ...snap,
        weights: DEFAULT_TRUST_WEIGHTS,
        components: {
          jobRating: { raw: snap.avgJobRating, weighted: (snap.avgJobRating / 5) * 100 * 0.4 },
          compliance: { raw: snap.complianceUptimePct, weighted: snap.complianceUptimePct * 0.25 },
          google: { raw: snap.googleRating, weighted: (snap.googleRating / 5) * 100 * 0.2 },
          experience: { raw: snap.experienceScore, weighted: snap.experienceScore * 0.15 },
        },
        reason: "seed_baseline",
        computedAt: ago(days(1)),
      },
    });
    await prisma.vendor.update({
      where: { id: snap.vendorId },
      data: { trustScore: snap.score, trustScoreAt: ago(days(1)) },
    });
  }

  // -------------------------------------------------------------------------
  // Chat threads (one per job+vendor, created once a bid exists)
  // -------------------------------------------------------------------------
  const thread1001 = await prisma.chatThread.create({
    data: {
      jobId: job1001.id,
      vendorId: vendors.loneStar.id,
      adminUnreadCount: 1,
      vendorUnreadCount: 0,
      lastMessageAt: ago(hours(3)),
      createdAt: ago(days(1)),
    },
  });
  const loneStarUser = await prisma.vendorUser.findFirstOrThrow({
    where: { vendorId: vendors.loneStar.id, isPrimary: true },
  });
  await prisma.chatMessage.createMany({
    data: [
      {
        threadId: thread1001.id,
        senderType: "ADMIN",
        senderAdminId: staff.id,
        body: "Rick — is the 4-day turn firm if the resident vacates on the 11th?",
        readByAdminAt: ago(days(1)),
        readByVendorAt: ago(days(1) - hours(1)),
        createdAt: ago(days(1)),
      },
      {
        threadId: thread1001.id,
        senderType: "VENDOR",
        senderVendorUserId: loneStarUser.id,
        body: "Firm as long as we get keys by noon. Paint crew is already blocked for that week.",
        readByVendorAt: ago(hours(3)),
        createdAt: ago(hours(3)),
      },
    ],
  });

  const thread1005 = await prisma.chatThread.create({
    data: {
      jobId: job1005.id,
      vendorId: vendors.bluebonnet.id,
      adminUnreadCount: 0,
      vendorUnreadCount: 2,
      lastMessageAt: ago(hours(9)),
      createdAt: ago(days(6)),
    },
  });
  await prisma.chatMessage.createMany({
    data: [
      {
        threadId: thread1005.id,
        senderType: "ADMIN",
        senderAdminId: manager.id,
        body: "Contract is out for signature — please countersign before the crew mobilizes Monday.",
        readByAdminAt: ago(hours(10)),
        createdAt: ago(hours(10)),
      },
      {
        threadId: thread1005.id,
        senderType: "SYSTEM",
        body: "Contract MOCK-CON-0088 sent for signature ($8,400.00).",
        readByAdminAt: ago(hours(9)),
        createdAt: ago(hours(9)),
      },
    ],
  });

  // -------------------------------------------------------------------------
  // Documents
  // -------------------------------------------------------------------------
  await prisma.jobDocument.createMany({
    data: [
      {
        jobId: job1005.id,
        vendorId: vendors.bluebonnet.id,
        type: "COI",
        fileName: "bluebonnet-coi-2026.pdf",
        url: "https://files.example.com/seed/bluebonnet-coi-2026.pdf",
        mimeType: "application/pdf",
        sizeBytes: 184320,
        uploadedByType: "VENDOR",
        uploadedById: vendors.bluebonnet.id,
        createdAt: ago(days(6)),
      },
      {
        jobId: job1005.id,
        type: "CONTRACT",
        fileName: "MOCK-CON-0088.pdf",
        url: "https://contracts.example.com/mock/MOCK-CON-0088",
        mimeType: "application/pdf",
        uploadedByType: "SYSTEM",
        createdAt: ago(days(2)),
      },
      {
        jobId: job1006.id,
        vendorId: vendors.loneStar.id,
        type: "PHOTO_AFTER",
        fileName: "unit-508-living-after.jpg",
        url: "https://files.example.com/seed/unit-508-living-after.jpg",
        mimeType: "image/jpeg",
        uploadedByType: "VENDOR",
        uploadedById: vendors.loneStar.id,
        createdAt: ago(days(9)),
      },
    ],
  });

  // -------------------------------------------------------------------------
  // Notification log (emergency dispatch trail)
  // -------------------------------------------------------------------------
  await prisma.notificationLog.createMany({
    data: [
      {
        channel: "SMS",
        status: "DELIVERED",
        recipientType: "VENDOR",
        vendorId: vendors.metroplexHvac.id,
        jobId: job1003.id,
        template: "emergency_dispatch",
        destination: "+19725551002",
        body: "EMERGENCY J-1003 — No cooling, Cortland Legacy Plano Unit 1122. Claim within 15 min.",
        providerKey: "twilio",
        providerMessageId: "SMseed0001",
        sentAt: ago(minutes(41)),
      },
      {
        channel: "SMS",
        status: "DELIVERED",
        recipientType: "VENDOR",
        vendorId: vendors.loneStar.id,
        jobId: job1003.id,
        template: "emergency_dispatch",
        destination: "+12145551001",
        body: "EMERGENCY J-1003 — No cooling, Cortland Legacy Plano Unit 1122. Claim within 15 min.",
        providerKey: "twilio",
        providerMessageId: "SMseed0002",
        sentAt: ago(minutes(41)),
      },
      {
        channel: "EMAIL",
        status: "SENT",
        recipientType: "ADMIN",
        adminUserId: staff.id,
        jobId: job1003.id,
        template: "emergency_unclaimed_escalation",
        destination: "priya.nair@cortland-example.com",
        body: "J-1003 went unclaimed past its 15-minute window. Radius widened to 65 miles.",
        providerKey: "smtp",
        sentAt: ago(minutes(26)),
      },
      {
        channel: "SMS",
        status: "DELIVERED",
        recipientType: "VENDOR",
        vendorId: vendors.rapidDry.id,
        jobId: job1004.id,
        template: "emergency_dispatch",
        destination: "+18175551003",
        body: "EMERGENCY J-1004 — Water intrusion, Cortland Arlington Commons Unit 308.",
        providerKey: "twilio",
        providerMessageId: "SMseed0003",
        sentAt: ago(hours(20)),
      },
    ],
  });

  // -------------------------------------------------------------------------
  // Activity log (rendered as the job timeline)
  // -------------------------------------------------------------------------
  await prisma.activityLog.createMany({
    data: [
      {
        entityType: "JOB",
        entityId: job1001.id,
        jobId: job1001.id,
        action: "job.created",
        toStatus: "OPEN",
        actorType: "ADMIN",
        actorAdminId: staff.id,
        summary: "Job J-1001 opened for bidding (cap enforced at $3,200.00).",
        createdAt: ago(days(2)),
      },
      {
        entityType: "BID",
        entityId: bid1001LoneStar.id,
        jobId: job1001.id,
        action: "bid.submitted",
        toStatus: "SUBMITTED",
        actorType: "VENDOR",
        actorVendorId: vendors.loneStar.id,
        summary: "Lone Star Make-Ready Co. submitted a bid of $2,680.00.",
        createdAt: ago(days(1)),
      },
      {
        entityType: "BID",
        entityId: bid1001LoneStar.id,
        jobId: job1001.id,
        action: "bid.withdrawn",
        fromStatus: "SUBMITTED",
        toStatus: "WITHDRAWN",
        actorType: "VENDOR",
        actorVendorId: vendors.rapidDry.id,
        summary: "RapidDry Water Restoration withdrew its bid.",
        createdAt: ago(hours(20)),
      },
      {
        entityType: "JOB",
        entityId: job1003.id,
        jobId: job1003.id,
        action: "emergency.dispatched",
        toStatus: "OPEN",
        actorType: "ADMIN",
        actorAdminId: staff.id,
        summary: "Emergency dispatched to 2 eligible vendors via SMS.",
        createdAt: ago(minutes(41)),
      },
      {
        entityType: "JOB",
        entityId: job1003.id,
        jobId: job1003.id,
        action: "emergency.escalated",
        actorType: "SYSTEM",
        summary: "Unclaimed after 15 minutes — radius widened to 65 miles, admin alerted.",
        metadata: { previousRadiusMiles: 40, newRadiusMiles: 65 },
        createdAt: ago(minutes(26)),
      },
      {
        entityType: "JOB",
        entityId: job1004.id,
        jobId: job1004.id,
        action: "emergency.claimed",
        fromStatus: "OPEN",
        toStatus: "AWARDED",
        actorType: "VENDOR",
        actorVendorId: vendors.rapidDry.id,
        summary: "RapidDry Water Restoration claimed the emergency in 6 minutes.",
        metadata: { responseMinutes: 6 },
        createdAt: ago(hours(20) - minutes(6)),
      },
      {
        entityType: "JOB",
        entityId: job1004.id,
        jobId: job1004.id,
        action: "emergency.on_site",
        actorType: "VENDOR",
        actorVendorId: vendors.rapidDry.id,
        summary: "Vendor marked on-site 32 minutes after dispatch.",
        metadata: { arrivalMinutes: 32 },
        createdAt: ago(hours(20) - minutes(32)),
      },
      {
        entityType: "BID",
        entityId: bid1005Bluebonnet.id,
        jobId: job1005.id,
        action: "bid.approved",
        fromStatus: "SUBMITTED",
        toStatus: "APPROVED",
        actorType: "ADMIN",
        actorAdminId: manager.id,
        summary: "Approved $8,400.00 — 12.5% above the $7,466.67 average of 3 bids.",
        metadata: { flagged: true, thresholdPct: 7.5 },
        createdAt: ago(days(2)),
      },
      {
        entityType: "CONTRACT",
        entityId: job1005.id,
        jobId: job1005.id,
        action: "contract.sent_for_signature",
        toStatus: "PENDING_SIGNATURE",
        actorType: "SYSTEM",
        summary: "GC job above $5,000.00 — contract MOCK-CON-0088 sent for signature.",
        createdAt: ago(days(2)),
      },
      {
        entityType: "RATING",
        entityId: job1006.id,
        jobId: job1006.id,
        action: "rating.submitted",
        actorType: "ADMIN",
        actorAdminId: staff.id,
        summary: "Rated 5/5 — no no-show, no missed deadline.",
        createdAt: ago(days(9)),
      },
    ],
  });

  // -------------------------------------------------------------------------
  console.log(`
Seed complete.

  Properties        ${await prisma.property.count()}
  Rent roll units   ${await prisma.rentRoll.count()}
  Vendors           ${await prisma.vendor.count()}
  Vendor logins     ${await prisma.vendorUser.count()}
  Jobs              ${await prisma.job.count()}  (${await prisma.job.count({ where: { priority: "EMERGENCY" } })} emergency)
  Bids              ${await prisma.bid.count()}
  Approval flags    ${await prisma.approvalFlag.count()}
  Activity entries  ${await prisma.activityLog.count()}

Admin logins (password: ${ADMIN_PASSWORD}) — /admin/login
  dana.reyes@cortland-example.com     OWNER
  marcus.hale@cortland-example.com    MANAGER
  priya.nair@cortland-example.com     STAFF

Vendor logins (password: ${VENDOR_PASSWORD}) — /vendor/login
  rick@lonestarmakeready.example.com  active, compliant, emergency-eligible
  sandra@metroplexhvac.example.com    active, compliant, emergency-eligible
  terrence@rapiddry.example.com       active, compliant, emergency-eligible
  marisol@bluebonnetgc.example.com    active, compliant, NOT emergency-eligible
  owen@trinityflooring.example.com    active, compliance expiring in 18 days
  dee@sunbeltelectric.example.com     active, compliance EXPIRED -> cannot bid
  hugo@ridgelinereno.example.com      active but out of radius (Houston)
  ava@pinnacleprops.example.com       PENDING approval -> no job access
  neil@gulfcoasthandyman.example.com  SUSPENDED -> cannot sign in
`);
}

main()
  .catch((err) => {
    console.error(err);
    process.exit(1);
  })
  .finally(async () => {
    await prisma.$disconnect();
  });
